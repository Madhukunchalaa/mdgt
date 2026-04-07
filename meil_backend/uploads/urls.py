from django.urls import path
from . import views

urlpatterns = [
    path('bulk-upload/', views.bulk_upload, name='bulk_upload'),
    path('bulk-upload', views.bulk_upload),
    path('get-fields/', views.get_model_fields, name='get_model_fields'),
    path('get-fields', views.get_model_fields),
    path('download-template/', views.generate_excel_template, name='generate_excel_template'),
    path('download-template', views.generate_excel_template),
]
