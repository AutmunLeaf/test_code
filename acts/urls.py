"""
URL-маршруты для приложения acts.
"""
from django.urls import path
from . import views

app_name = 'acts'

urlpatterns = [
    # Главная
    path('', views.index, name='index'),
    
    # Создание
    path('ks2/create/', views.create_act, {'act_type': 'ks2'}, name='ks2_create'),
    path('ks3/create/', views.create_act, {'act_type': 'ks3'}, name='ks3_create'),
    
    # Просмотр / Редактирование
    path('<uuid:pk>/', views.act_detail, name='detail'),
    path('<uuid:pk>/edit/', views.edit_act, name='edit'),
    
    # Скачивание
    path('<uuid:pk>/download/<str:format>/', views.download_act, name='download'),
    
    # Удаление
    path('<uuid:pk>/delete/', views.delete_act, name='delete'),
    
    # API
    path('api/add-work/<str:act_type>/', views.api_add_work, name='api_add_work'),
    
    path('<uuid:pk>/status/', views.update_act_status, name='update_status'),
]