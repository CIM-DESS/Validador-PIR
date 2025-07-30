from django.urls import path
from . import views
from django.contrib.auth import views as auth_views

urlpatterns = [
    path('', views.inicio, name='inicio'),
    path('descargar_excel/', views.descargar_excel, name='descargar_excel'),
    path('estado/<str:tarea_id>/', views.estado_procesamiento, name='estado_procesamiento'),  # � Nueva ruta para estado
    path('accounts/login/', auth_views.LoginView.as_view(), name='login'),
    path('accounts/logout/', auth_views.LogoutView.as_view(), name='logout'),
]
