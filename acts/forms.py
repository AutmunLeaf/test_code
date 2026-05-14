"""
Формы для ввода данных актов.
"""
from django import forms

from .constants import MAX_WORK_ROWS_KS2, MAX_WORK_ROWS_KS3
from django.forms import inlineformset_factory, BaseInlineFormSet
from .models import ActInput, WorkItem


class DateInput(forms.DateInput):
    """Виджет для input type="date" с правильным форматом"""
    input_type = 'date'
    
    def __init__(self, attrs=None, format=None):
        # Формат для отображения в HTML (обязательно YYYY-MM-DD!)
        super().__init__(attrs, format='%Y-%m-%d')
        if attrs is None:
            attrs = {}
        attrs.setdefault('class', 'form-control')
        self.attrs = attrs


class ActInputForm(forms.ModelForm):
    """Основная форма для ввода шапки акта"""

    contract_date = forms.DateField(
        widget=DateInput(),
        label='Дата договора',
        input_formats=['%Y-%m-%d', '%d.%m.%Y']  # ✅ Принимаем оба формата
    )
    report_from = forms.DateField(
        widget=DateInput(),
        label='Период с',
        input_formats=['%Y-%m-%d', '%d.%m.%Y']
    )
    report_to = forms.DateField(
        widget=DateInput(),
        label='Период по',
        input_formats=['%Y-%m-%d', '%d.%m.%Y']
    )
    
    class Meta:
        model = ActInput
        fields = [
            'document_number', 'contract_number', 'contract_date',
            'report_from', 'report_to',
            'investor', 'investor_okpo', 'customer', 'customer_okpo',
            'contractor', 'contractor_okpo', 'okdp',
            'construction', 'object_name', 'operation',
            'surrender_position', 'surrender_signature',
            'accept_position', 'accept_signature',
            'vat_rate', 'smeta',
        ]
        labels = {
            'document_number': 'Номер акта',
            'contract_number': 'Номер договора',
            'investor': 'Инвестор (необязательно)',
            'investor_okpo': 'ОКПО инвестора',
            'contractor': 'Подрядчик',
            'contractor_okpo': 'ОКПО подрядчика',
            'okdp': 'ОКДП',
            'customer': 'Заказчик / Генподрядчик',
            'construction': 'Наименование стройки',
            'object_name': 'Объект (если есть)',
            'operation': 'Вид операции',
            'surrender_position': 'Должность сдавшего',
            'surrender_signature': 'ФИО сдавшего',
            'accept_position': 'Должность принявшего',
            'accept_signature': 'ФИО принявшего',
            'vat_rate': 'Ставка НДС',
            'smeta': 'Сметная стоимость',
        }
        widgets = {
            'contractor': forms.TextInput(attrs={'class': 'form-control', 'placeholder': 'Наименование подрядчика'}),
            'contractor_okpo': forms.TextInput(attrs={'class': 'form-control', 'placeholder': 'ОКПО'}),
            'okdp': forms.TextInput(attrs={'class': 'form-control', 'placeholder': 'ОКДП'}),

            'investor': forms.TextInput(attrs={'class': 'form-control', 'placeholder': 'При отсутствии — оставьте пустым'}),
            'customer': forms.TextInput(attrs={'class': 'form-control', 'placeholder': 'Введите наименование'}),
            'construction': forms.TextInput(attrs={'class': 'form-control', 'placeholder': 'Например: ЖК "Солнечный", г. Сургут'}),
            'object_name': forms.TextInput(attrs={'class': 'form-control', 'placeholder': 'Например: Корпус 1, Секция А...'}),
            'operation': forms.TextInput(attrs={'class': 'form-control'}),
            'surrender_position': forms.TextInput(attrs={'class': 'form-control'}),
            'surrender_signature': forms.TextInput(attrs={'class': 'form-control'}),
            'accept_position': forms.TextInput(attrs={'class': 'form-control'}),
            'accept_signature': forms.TextInput(attrs={'class': 'form-control'}),
            'vat_rate': forms.NumberInput(attrs={
                'class': 'form-control',
                'placeholder': 'Например: 22, 20, 10, 0',
                'step': '1',
                'min': '0',
                'max': '100'
            }),
            'smeta': forms.TextInput(attrs={'class': 'form-control', 'placeholder': 'Введите значение'}),
            'status': forms.Select(attrs={'class': 'form-select'}),
        }
    
    def __init__(self, *args, act_type=None, **kwargs):
        self._act_type = act_type
        super().__init__(*args, **kwargs)
        at = act_type or self.initial.get('act_type') or getattr(self.instance, 'act_type', None)

        if at == 'ks3':
            self.fields['object_name'].widget = forms.HiddenInput()
            self.fields['smeta'].widget = forms.HiddenInput()
        elif at == 'ks2':
            self.fields['operation'].widget = forms.HiddenInput()

    def clean(self):
        cleaned_data = super().clean()
        report_from = cleaned_data.get('report_from')
        report_to = cleaned_data.get('report_to')

        if report_from and report_to and report_to <= report_from:
            raise forms.ValidationError(
                '❌ Дата окончания периода должна быть позже даты начала.'
            )
        contractor = (cleaned_data.get('contractor') or '').strip()
        if not contractor:
            self.add_error('contractor', 'Укажите подрядчика.')
        return cleaned_data

class KS2WorkForm(forms.ModelForm):
    """Форма для работы в КС-2"""
    class Meta:
        model = WorkItem
        fields = ['position', 'name', 'number_pricelist', 'unit', 'quantity', 'price', 'order']
        labels = {
            'position': 'Поз.', 'name': 'Наименование работ', 'number_pricelist': '№ расценки',
            'unit': 'Ед.', 'quantity': 'Кол-во', 'price': 'Цена, руб.', 'order': 'Порядок',
        }
        widgets = {
            'name': forms.Textarea(attrs={'rows': 2, 'class': 'form-control'}),
            'position': forms.TextInput(attrs={'class': 'form-control', 'placeholder': '20-30'}),
            'number_pricelist': forms.TextInput(attrs={'class': 'form-control'}),
            'unit': forms.TextInput(attrs={'class': 'form-control', 'placeholder': 'м2, шт, т...'}),
            'quantity': forms.NumberInput(attrs={'class': 'form-control'}),
            'price': forms.NumberInput(attrs={'class': 'form-control'}),
            'order': forms.NumberInput(attrs={'class': 'form-control', 'value': '0'}),
        }


class KS3WorkForm(forms.ModelForm):
    """Форма для работы в КС-3"""
    class Meta:
        model = WorkItem
        fields = ['name', 'code', 'cost_from_start', 'cost_from_year', 'cost_report_period', 'order']
        labels = {
            'name': 'Наименование работ/затрат', 'code': 'Код',
            'cost_from_start': 'С начала работ, руб.', 'cost_from_year': 'С начала года, руб.',
            'cost_report_period': 'За период, руб.', 'order': 'Порядок',
        }
        widgets = {
            'name': forms.Textarea(attrs={'rows': 2, 'class': 'form-control'}),
            'code': forms.TextInput(attrs={'class': 'form-control', 'placeholder': '20-30'}),
            'cost_from_start': forms.NumberInput(attrs={'class': 'form-control'}),
            'cost_from_year': forms.NumberInput(attrs={'class': 'form-control'}),
            'cost_report_period': forms.NumberInput(attrs={'class': 'form-control'}),
            'order': forms.NumberInput(attrs={'class': 'form-control', 'value': '0'}),
        }


# === ФАБРИКИ FORMSET (возвращают КЛАССЫ) ===
def make_ks2_work_formset(extra: int = 1):
    """extra — число пустых/начальных строк, нужно для импорта КС-6а."""
    extra = min(max(int(extra), 1), MAX_WORK_ROWS_KS2)
    return inlineformset_factory(
        ActInput, WorkItem,
        form=KS2WorkForm,
        extra=extra,
        can_delete=True,
        can_order=True,
        max_num=MAX_WORK_ROWS_KS2,
        validate_max=True,
    )


def make_ks3_work_formset(extra: int = 1):
    extra = min(max(int(extra), 1), MAX_WORK_ROWS_KS3)
    return inlineformset_factory(
        ActInput, WorkItem,
        form=KS3WorkForm,
        extra=extra,
        can_delete=True,
        can_order=True,
        max_num=MAX_WORK_ROWS_KS3,
        validate_max=True,
    )


def make_work_formset_class(act_type: str, extra: int = 1):
    if act_type == 'ks2':
        return make_ks2_work_formset(extra)
    if act_type == 'ks3':
        return make_ks3_work_formset(extra)
    raise ValueError(f'Неизвестный тип акта: {act_type}')


KS2WorkFormSet = make_ks2_work_formset(1)
KS3WorkFormSet = make_ks3_work_formset(1)