"""
Представления (views) для приложения актов.
"""
from django.shortcuts import render, redirect, get_object_or_404
from django.contrib.auth.decorators import login_required
from django.contrib import messages
from django.http import HttpResponse, JsonResponse
from django.conf import settings
from django.db import transaction
import os
import tempfile
import time
import logging

from .models import ActInput, WorkItem, OrgConstants
from .forms import (
    ActInputForm, KS2WorkFormSet, KS3WorkFormSet,
    KS2WorkForm, KS3WorkForm
)
from .utils.ks2_gen import fill_ks2
from .utils.ks3_gen import fill_ks3

logger = logging.getLogger(__name__)


def log_action(user, action, target_model, target_id, details=None):
    """
    Вспомогательная функция для записи в журнал аудита.
    Если у тебя есть модель AuditLog — раскомментируй код ниже.
    """
    # from .models import AuditLog
    # AuditLog.objects.create(
    #     user=user, action=action, target_model=target_model,
    #     target_id=str(target_id), details=details or {}
    # )
    logger.info(f"[AUDIT] {user} → {action} {target_model}#{target_id} | {details}")


@login_required
def index(request):
    """Главная страница — список актов"""
    acts = ActInput.objects.select_related('created_by').all()[:50]
    
    ks2_acts = [a for a in acts if a.act_type == 'ks2']
    ks3_acts = [a for a in acts if a.act_type == 'ks3']
    
    return render(request, 'acts/index.html', {
        'ks2_acts': ks2_acts,
        'ks3_acts': ks3_acts,
    })


@login_required
def create_act(request, act_type):
    """Создание нового акта (КС-2 или КС-3)"""
    if act_type not in ['ks2', 'ks3']:
        messages.error(request, 'Неверный тип акта')
        return redirect('acts:index')
    
    FormSetFactory = KS2WorkFormSet if act_type == 'ks2' else KS3WorkFormSet
    
    if request.method == 'POST':
        form = ActInputForm(request.POST)
        formset = FormSetFactory(request.POST, prefix='works')

        if not form.is_valid():
            logger.warning(f"❌ Form errors: {form.errors}")
        if not formset.is_valid():
            logger.warning(f"❌ Formset errors: {formset.errors}")
        
        if form.is_valid() and formset.is_valid():
            with transaction.atomic():
                act = form.save(commit=False)
                act.act_type = act_type
                act.created_by = request.user
                act.save()
                
                works = formset.save(commit=False)
                for work in works:
                    work.act = act
                    work.save()
                
                for obj in formset.deleted_objects:
                    obj.delete()

                log_action(request.user, 'create', 'ActInput', act.id, {'type': act_type})
                
                messages.success(
                    request, 
                    f'✅ {act.get_act_type_display()} №{act.document_number} создан!'
                )
                return redirect('acts:download', pk=act.pk, format='pdf')
    else:
        form = ActInputForm(initial={
            'act_type': act_type,
            'contractor': OrgConstants.NAME,
            'contractor_okpo': OrgConstants.OKPO,
            'okdp': OrgConstants.OKDP,
        })
        formset = FormSetFactory(prefix='works')
    
    return render(request, 'acts/act_form.html', {
        'form': form,
        'formset': formset,
        'act_type': act_type,
        'act_type_display': 'КС-2' if act_type == 'ks2' else 'КС-3',
    })


@login_required
def edit_act(request, pk):
    """Редактирование существующего акта"""
    act = get_object_or_404(ActInput, pk=pk)
    
    FormSet = KS2WorkFormSet if act.act_type == 'ks2' else KS3WorkFormSet
    
    if request.method == 'POST':
        form = ActInputForm(request.POST, instance=act)
        formset = FormSet(request.POST, instance=act, prefix='works')
        
        if form.is_valid() and formset.is_valid():
            with transaction.atomic():
                act = form.save()
                
                instances = formset.save(commit=False)
                for instance in instances:
                    instance.act = act
                    instance.save()
                
                for obj in formset.deleted_objects:
                    obj.delete()
                
                log_action(request.user, 'update', 'ActInput', act.id)
            
            messages.success(request, f'✅ Акт №{act.document_number} обновлён!')
            return redirect('acts:detail', pk=act.pk)
        else:
            if not form.is_valid():
                logger.warning(f"❌ Form errors: {form.errors}")
            if not formset.is_valid():
                logger.warning(f"❌ Formset errors: {formset.errors}")
    else:
        form = ActInputForm(instance=act)
        formset = FormSet(instance=act, prefix='works')
    
    return render(request, 'acts/act_form.html', {
        'form': form,
        'formset': formset,
        'act_type': act.act_type,
        'act_type_display': act.get_act_type_display(),
        'act': act,
    })


@login_required
def act_detail(request, pk):
    """Просмотр акта с кнопками скачивания"""
    act = get_object_or_404(ActInput.objects.prefetch_related('works'), pk=pk)
    return render(request, 'acts/act_detail.html', {
        'act': act,
        'statuses': ActInput.STATUSES  # 👈 ActInput, а не Act
    })


@login_required
def update_act_status(request, pk):
    """Быстрая смена статуса акта"""
    act = get_object_or_404(ActInput, pk=pk)  # 👈 ActInput
    if request.method == 'POST':
        new_status = request.POST.get('status')
        valid_statuses = dict(ActInput.STATUSES)  # 👈 ActInput
        
        if new_status in valid_statuses:
            old_status = act.status
            act.status = new_status
            act.save(update_fields=['status'])
            
            log_action(request.user, 'update_status', 'ActInput', act.id, {
                'old': old_status, 'new': new_status
            })
            messages.success(request, f'✅ Статус изменён на «{valid_statuses[new_status]}»')
        else:
            messages.error(request, '❌ Выбран недопустимый статус')
            
    return redirect('acts:detail', pk=act.id)


@login_required
def download_act(request, pk, format='pdf'):
    if format not in ['pdf', 'xlsx']:
        return HttpResponse('Неверный формат', status=400)

    act = get_object_or_404(ActInput, pk=pk)
    act.refresh_from_db()

    works_qs = WorkItem.objects.filter(act=act).order_by('order', 'id')

    logger.info(f"📥 Генерация {format} для акта {act.document_number}")
    for i, w in enumerate(works_qs):
        if act.act_type == 'ks2':
            logger.info(f"   Работа {i+1}: {w.name} | Кол-во: {w.quantity} | Цена: {w.price}")
        else:
            logger.info(f"   Работа {i+1}: {w.name} | За период: {w.cost_report_period}")

    template_name = f'ks{act.act_type[2]}.xlsx'
    template_path = os.path.join(settings.MEDIA_ROOT, 'templates', template_name)
    if not os.path.exists(template_path):
        messages.error(request, f'❌ Шаблон {act.get_template_name()} не найден!')
        return redirect('acts:detail', pk=act.pk)

    data = {
        "investor": act.investor, "okpo_investor": act.investor_okpo,
        "customer": act.customer, "okpo_customer": act.customer_okpo,
        "contractor": act.contractor, "okpo_contractor": act.contractor_okpo,
        "construction": act.construction, "object": act.object_name,
        "operation": act.operation, "okdp": act.okdp,
        "document_number": act.document_number,
        "contract_number": act.contract_number,
        "contract_date": act.contract_date.strftime('%d.%m.%Y'),
        "day_contract": act.contract_date.day,
        "month_contract": act.contract_date.month,
        "year_contract": act.contract_date.year,
        "report_from": act.report_from.strftime('%d.%m.%Y'),
        "report_to": act.report_to.strftime('%d.%m.%Y'),
        "surrender_position": act.surrender_position,
        "surrender_signature": act.surrender_signature,
        "accept_position": act.accept_position,
        "accept_signature": act.accept_signature,
        "vat_rate": act.vat_rate,
        "smeta": act.smeta,
        "works": [
            {
                "position": w.position, "name": w.name, "code": w.code,
                "number_pricelist": w.number_pricelist,
                "unit_of_measurement": w.unit,
                "quantity": float(w.quantity or 0),
                "price": float(w.price or 0),
                "cost_from_start": float(w.cost_from_start or 0),
                "cost_from_year": float(w.cost_from_year or 0),
                "cost_report_period": float(w.cost_report_period or 0),
            }
            for w in works_qs
        ]
    }

    output_path = None
    try:
        output_dir = settings.MEDIA_ROOT / 'output'
        output_dir.mkdir(parents=True, exist_ok=True)
        
        with tempfile.NamedTemporaryFile(delete=False, suffix=f".{format}", dir=output_dir) as tmp:
            output_path = tmp.name

        if act.act_type == 'ks2':
            fill_ks2(template_path, output_path, data, format=format)
        else:
            fill_ks3(template_path, output_path, data, format=format)

        with open(output_path, 'rb') as f:
            file_content = f.read()

        timestamp = int(time.time())
        filename = f"{act.act_type.upper()}_{act.document_number}_{timestamp}.{format}"
        content_type = 'application/pdf' if format == 'pdf' else 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'

        response = HttpResponse(file_content, content_type=content_type)
        response['Content-Disposition'] = f'attachment; filename="{filename}"'
        response['Cache-Control'] = 'no-cache, no-store, must-revalidate'
        return response

    except Exception as e:
        logger.error(f"❌ Ошибка генерации: {e}", exc_info=True)
        messages.error(request, f'❌ Ошибка генерации: {e}')
        return redirect('acts:detail', pk=act.pk)
        
    finally:
        if output_path and os.path.exists(output_path):
            try:
                os.unlink(output_path)
            except PermissionError:
                pass


@login_required
def delete_act(request, pk):
    """Удаление акта (только черновики)"""
    act = get_object_or_404(ActInput, pk=pk)
    
    if act.status != 'draft':
        messages.error(request, '❌ Можно удалять только черновики')
        return redirect('acts:detail', pk=act.pk)
    
    if request.method == 'POST':
        doc_num = act.document_number
        log_action(request.user, 'delete', 'ActInput', act.id)
        act.delete()
        messages.success(request, f'🗑️ Акт №{doc_num} удалён')
        return redirect('acts:index')
    
    return render(request, 'acts/act_confirm_delete.html', {'act': act})


@login_required
def api_add_work(request, act_type):
    """Возвращает пустую форму работы для JS-добавления"""
    if act_type == 'ks2':
        form = KS2WorkForm(prefix='works-__prefix__')
    elif act_type == 'ks3':
        form = KS3WorkForm(prefix='works-__prefix__')
    else:
        return JsonResponse({'error': 'Неверный тип'}, status=400)
    
    return JsonResponse({'html': form.as_p()})
