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
from pathlib import Path

from .models import ActInput, WorkItem
from .forms import (
    ActInputForm,
    KS2WorkForm,
    KS3WorkForm,
    make_work_formset_class,
)
from .utils.ks2_gen import fill_ks2
from .utils.ks3_gen import fill_ks3
from .utils.ks6a_import import parse_ks6a_workbook

logger = logging.getLogger(__name__)

FORMSET_PREFIX = 'works'


def _work_formset_extra_from_post(post_data):
    try:
        n = int(post_data.get(f'{FORMSET_PREFIX}-TOTAL_FORMS', 1))
    except (TypeError, ValueError):
        n = 1
    return min(max(n, 1), 14)


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
def import_ks6a(request):
    """Загрузка Excel КС-6а: строки работ в сессию и редирект на форму акта."""
    if request.method != 'POST':
        return redirect('acts:index')

    act_type = request.POST.get('act_type')
    if act_type not in ('ks2', 'ks3'):
        messages.error(request, 'Неверный тип акта')
        return redirect('acts:index')

    upload = request.FILES.get('ks6a_file')
    if not upload:
        messages.error(request, 'Выберите файл КС-6а (Excel).')
        return redirect('acts:ks2_create' if act_type == 'ks2' else 'acts:ks3_create')

    ext = Path(upload.name).suffix.lower()
    if ext not in ('.xlsx', '.xls'):
        messages.error(request, 'Нужен файл Excel с расширением .xlsx или .xls.')
        return redirect('acts:ks2_create' if act_type == 'ks2' else 'acts:ks3_create')

    act_pk = request.POST.get('act_pk')
    tmp_path = None
    try:
        with tempfile.NamedTemporaryFile(delete=False, suffix=ext) as tmp:
            tmp_path = tmp.name
            for chunk in upload.chunks():
                tmp.write(chunk)
        works, act_initial = parse_ks6a_workbook(tmp_path)
    except Exception as e:
        logger.exception('Импорт КС-6а')
        messages.error(request, f'Не удалось прочитать файл: {e}')
        if act_pk:
            return redirect('acts:edit', pk=act_pk)
        return redirect('acts:ks2_create' if act_type == 'ks2' else 'acts:ks3_create')
    finally:
        if tmp_path and os.path.exists(tmp_path):
            try:
                os.unlink(tmp_path)
            except OSError:
                pass

    has_header = any((v or '').strip() for v in (act_initial or {}).values())
    if not works and not has_header:
        messages.warning(
            request,
            'Не найдено строк работ и реквизитов из фиксированных ячеек. Проверьте первый лист КС-6а.'
        )
        if act_pk:
            return redirect('acts:edit', pk=act_pk)
        return redirect('acts:ks2_create' if act_type == 'ks2' else 'acts:ks3_create')

    total = len(works)
    if total > 14:
        messages.warning(
            request,
            f'В журнале найдено {total} строк; в форму перенесены только первые 14 (лимит акта).',
        )
        works = works[:14]

    request.session['ks6a_import'] = {
        'works': works,
        'act_type': act_type,
        'act_initial': act_initial or {},
    }
    parts = []
    if works:
        parts.append(f'строк работ: {len(works)}')
    if has_header:
        parts.append('реквизиты из ячеек шапки')
    messages.success(request, 'Из КС-6а подставлено: ' + ', '.join(parts) + '.')

    if act_pk:
        act = get_object_or_404(ActInput, pk=act_pk)
        if act.act_type != act_type:
            messages.error(request, 'Тип акта не совпадает с выбранной формой импорта.')
            return redirect('acts:edit', pk=act.pk)
        if act.status != 'draft':
            messages.error(request, 'Импорт доступен только для актов в статусе «Черновик».')
            return redirect('acts:edit', pk=act.pk)
        with transaction.atomic():
            WorkItem.objects.filter(act=act).delete()
        return redirect('acts:edit', pk=act.pk)

    return redirect('acts:ks2_create' if act_type == 'ks2' else 'acts:ks3_create')


@login_required
def create_act(request, act_type):
    """Создание нового акта (КС-2 или КС-3)"""
    if act_type not in ['ks2', 'ks3']:
        messages.error(request, 'Неверный тип акта')
        return redirect('acts:index')

    if request.method == 'POST':
        extra = _work_formset_extra_from_post(request.POST)
        FormSetFactory = make_work_formset_class(act_type, extra)
        form = ActInputForm(request.POST, act_type=act_type)
        formset = FormSetFactory(request.POST, prefix=FORMSET_PREFIX)

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
        form_initial = {}
        imp = request.session.pop('ks6a_import', None)
        if imp and imp.get('act_type') == act_type:
            ai = {k: v for k, v in (imp.get('act_initial') or {}).items() if (v or '').strip()}
            form_initial = {**form_initial, **ai}
            works = imp.get('works') or []
            n = min(len(works), 14)
            FormSetFactory = make_work_formset_class(act_type, max(n, 1))
            formset = FormSetFactory(initial=works[:14], prefix=FORMSET_PREFIX)
        else:
            FormSetFactory = make_work_formset_class(act_type, 1)
            formset = FormSetFactory(prefix=FORMSET_PREFIX)
        form = ActInputForm(initial=form_initial, act_type=act_type)

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

    if request.method == 'POST':
        extra = _work_formset_extra_from_post(request.POST)
        FormSet = make_work_formset_class(act.act_type, extra)
        form = ActInputForm(request.POST, instance=act, act_type=act.act_type)
        formset = FormSet(request.POST, instance=act, prefix=FORMSET_PREFIX)
        
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
        imp = request.session.pop('ks6a_import', None)
        if imp and imp.get('act_type') == act.act_type:
            ai = {k: v for k, v in (imp.get('act_initial') or {}).items() if (v or '').strip()}
            form = ActInputForm(instance=act, initial=ai, act_type=act.act_type)
            works = imp.get('works') or []
            n = min(len(works), 14)
            FormSet = make_work_formset_class(act.act_type, max(n, 1))
            formset = FormSet(initial=works[:14], instance=act, prefix=FORMSET_PREFIX)
        else:
            form = ActInputForm(instance=act, act_type=act.act_type)
            FormSet = make_work_formset_class(act.act_type, 1)
            formset = FormSet(instance=act, prefix=FORMSET_PREFIX)

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
