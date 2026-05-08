"""
Админка для приложения актов.
"""
from django.contrib import admin
from django.utils.html import format_html
from .models import ActInput, WorkItem, OrgConstants


@admin.register(ActInput)
class ActInputAdmin(admin.ModelAdmin):
    list_display = [
        'document_number', 'act_type', 'contract_number', 
        'report_period', 'construction', 'status', 'created_at'
    ]
    list_filter = ['act_type', 'status', 'created_at']
    search_fields = ['document_number', 'contract_number', 'construction', 'investor', 'customer']
    readonly_fields = ['created_at', 'updated_at', 'created_by', 'contractor', 'contractor_okpo', 'okdp']
    date_hierarchy = 'created_at'
    
    fieldsets = (
        ('Основное', {
            'fields': ('act_type', 'document_number', 'status')
        }),
        ('Договор и период', {
            'fields': ('contract_number', 'contract_date', 'report_from', 'report_to')
        }),
        ('Участники', {
            'fields': (
                ('investor', 'investor_okpo'),
                ('customer', 'customer_okpo'),
                ('contractor', 'contractor_okpo'),
            )
        }),
        ('Объект', {
            'fields': ('construction', 'object_name', 'operation', 'okdp')
        }),
        ('Подписи', {
            'fields': (
                ('surrender_position', 'surrender_signature'),
                ('accept_position', 'accept_signature'),
            )
        }),
        ('Финансы', {
            'fields': ('vat_rate', 'smeta')
        }),
        ('Системные', {
            'fields': ('created_by', 'created_at', 'updated_at'),
            'classes': ('collapse',)
        }),
    )
    
    def report_period(self, obj):
        return f"{obj.report_from} — {obj.report_to}"
    report_period.short_description = 'Период'
    
    def save_model(self, request, obj, form, change):
        if not obj.pk:
            obj.created_by = request.user
        super().save_model(request, obj, form, change)


class WorkItemInline(admin.TabularInline):
    model = WorkItem
    extra = 0
    can_delete = True
    
    def has_add_permission(self, request, obj=None):
        return False  # Работы добавляем через веб-форму, не в админке


@admin.register(WorkItem)
class WorkItemAdmin(admin.ModelAdmin):
    list_display = ['act', 'position', 'name_short', 'order']
    list_filter = ['act__act_type', 'act__status']
    search_fields = ['name', 'position', 'code']
    
    def name_short(self, obj):
        return obj.name[:50] + '...' if len(obj.name) > 50 else obj.name
    name_short.short_description = 'Наименование'


# Регистрируем константы как инфо-панель
class OrgConstantsAdmin(admin.ModelAdmin):
    def changelist_view(self, request, extra_context=None):
        extra_context = extra_context or {}
        extra_context['constants'] = OrgConstants
        return super().changelist_view(request, extra_context)
    
    def has_add_permission(self, request):
        return False
    
    def has_delete_permission(self, request, obj=None):
        return False


# Показываем реквизиты в админке
admin.site.site_header = 'Администрирование — Мостоотряд-69'
admin.site.site_title = 'Мостоотряд-69 | Админка'
admin.site.index_title = 'Панель управления актами КС-2/КС-3'