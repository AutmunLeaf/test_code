"""
Модели для приложения актов КС-2 и КС-3.
"""
from django.db import models
from django.contrib.auth.models import User
import uuid


class OrgConstants:
    """
    Статические реквизиты организации Мостоотряд-69.
    Править только здесь — подставятся везде автоматически.
    Источник: [[1]], [[9]]
    """
    NAME = 'ООО "МОСТООТРЯД-69"'
    INN = '8603236064'
    KPP = '770301001'
    OKPO = '34810147'  # [[1]]
    OGRN = '1188617017247'
    OKDP = '4530'  # Строительство гидротехнических сооружений [[9]]
    ADDRESS = '628400, ХМАО-Югра, г. Сургут, ул. Мира, д. 69'
    PHONE = '+7 (3462) 12-34-56'
    EMAIL = 'info@mo69.ru'
    DIRECTOR = 'Иванов Иван Иванович'
    DIRECTOR_POSITION = 'Генеральный директор'


class ActInput(models.Model):
    """
    Универсальная модель для хранения введённых данных актов КС-2 и КС-3.
    """
    ACT_TYPE_CHOICES = [
        ('ks2', 'Акт КС-2 (приёмка работ)'),
        ('ks3', 'Справка КС-3 (стоимость работ)'),
    ]
    
    id = models.UUIDField(primary_key=True, default=uuid.uuid4, editable=False)
    created_at = models.DateTimeField(auto_now_add=True, verbose_name='Дата создания')
    updated_at = models.DateTimeField(auto_now=True, verbose_name='Дата обновления')
    created_by = models.ForeignKey(
        User, on_delete=models.SET_NULL, null=True, blank=True,
        verbose_name='Создал пользователь'
    )
    
    # Тип акта
    act_type = models.CharField(
        max_length=3, choices=ACT_TYPE_CHOICES,
        verbose_name='Тип акта'
    )
    
    # === Шапка документа ===
    document_number = models.CharField(max_length=50, verbose_name='Номер документа')
    contract_number = models.CharField(max_length=50, verbose_name='Номер договора')
    contract_date = models.DateField(verbose_name='Дата договора')
    report_from = models.DateField(verbose_name='Отчётный период с')
    report_to = models.DateField(verbose_name='Отчётный период по')
    
    # === Участники ===
    investor = models.CharField(max_length=200, verbose_name='Инвестор (наименование)')
    investor_okpo = models.CharField(max_length=20, verbose_name='ОКПО инвестора')
    
    customer = models.CharField(max_length=200, verbose_name='Заказчик/Генподрядчик')
    customer_okpo = models.CharField(max_length=20, verbose_name='ОКПО заказчика')
    
    # Подрядчик — всегда Мостоотряд-69 (статика)
    contractor = models.CharField(
        max_length=200, default=OrgConstants.NAME,
        verbose_name='Подрядчик', editable=False
    )
    contractor_okpo = models.CharField(
        max_length=20, default=OrgConstants.OKPO,
        verbose_name='ОКПО подрядчика', editable=False
    )
    
    # === Объект строительства ===
    construction = models.CharField(max_length=200, verbose_name='Стройка (наименование, адрес)')
    object_name = models.CharField(max_length=200, blank=True, verbose_name='Объект (для КС-2)')
    operation = models.CharField(
        max_length=30, blank=True, default='',
        verbose_name='Вид операции (для КС-3)'
    )
    
    # === Вид деятельности ===
    okdp = models.CharField(
        max_length=10, default=OrgConstants.OKDP,
        verbose_name='Вид деятельности по ОКДП', editable=False
    )
    
    # === Подписи ===
    surrender_position = models.CharField(max_length=100, verbose_name='Должность сдавшего')
    surrender_signature = models.CharField(max_length=100, verbose_name='Подпись сдавшего (ФИО)')
    
    accept_position = models.CharField(max_length=100, verbose_name='Должность принявшего')
    accept_signature = models.CharField(max_length=100, verbose_name='Подпись принявшего (ФИО)')
    
    # === Финансы ===
    vat_rate = models.CharField(max_length=10, default='20%', verbose_name='Ставка НДС')
    smeta = models.CharField(max_length=50, blank=True, verbose_name='Сметная стоимость (для КС-2)')
    
    # Статус
    STATUSES = [
        ('draft', 'Черновик'),
        ('sent', 'Отправлен'),
        ('approved', 'Согласован'),
        ('rejected', 'Отклонён'),
    ]
    status = models.CharField(
        max_length=10, choices=STATUSES, default='draft',
        verbose_name='Статус'
    )
    
    class Meta:
        db_table = 'acts_input'
        ordering = ['-created_at']
        verbose_name = 'Ввод данных акта'
        verbose_name_plural = 'Ввод данных актов'
    
    def __str__(self):
        return f"{self.get_act_type_display()} №{self.document_number} от {self.report_to}"
    
    def get_template_name(self):
        """Возвращает имя шаблона Excel"""
        return f"{self.act_type}.xlsx"


class WorkItem(models.Model):
    """
    Элемент работы/услуги для акта.
    Используется и для КС-2, и для КС-3 (разные поля).
    """
    act = models.ForeignKey(
        ActInput, on_delete=models.CASCADE, related_name='works',
        verbose_name='Акт'
    )
    
    # === Общие поля ===
    position = models.CharField(max_length=20, blank=True, verbose_name='Позиция по смете')
    name = models.TextField(verbose_name='Наименование работ/затрат')
    code = models.CharField(max_length=20, blank=True, verbose_name='Код (для КС-3)')
    number_pricelist = models.CharField(max_length=50, blank=True, verbose_name='№ расценки (для КС-2)')
    
    # === Для КС-2 ===
    unit = models.CharField(max_length=20, blank=True, verbose_name='Ед. измерения')
    quantity = models.DecimalField(
        max_digits=12, decimal_places=3, null=True, blank=True,
        verbose_name='Количество'
    )
    price = models.DecimalField(
        max_digits=15, decimal_places=2, null=True, blank=True,
        verbose_name='Цена за единицу, руб.'
    )
    
    # === Для КС-3 ===
    cost_from_start = models.DecimalField(
        max_digits=15, decimal_places=2, default=0,
        verbose_name='Стоимость с начала проведения работ, руб.'
    )
    cost_from_year = models.DecimalField(
        max_digits=15, decimal_places=2, default=0,
        verbose_name='Стоимость с начала года, руб.'
    )
    cost_report_period = models.DecimalField(
        max_digits=15, decimal_places=2, default=0,
        verbose_name='Стоимость за отчётный период, руб.'
    )
    
    # Порядок сортировки
    order = models.PositiveIntegerField(default=0, verbose_name='Порядковый номер')
    
    class Meta:
        db_table = 'acts_work_items'
        ordering = ['order', 'id']
        verbose_name = 'Вид работ'
        verbose_name_plural = 'Виды работ'
    
    def __str__(self):
        return f"{self.position or '#'} — {self.name[:50]}..."
    
    @property
    def total_cost_ks2(self):
        """Расчёт стоимости для КС-2: количество × цена"""
        if self.quantity and self.price:
            return self.quantity * self.price
        return 0