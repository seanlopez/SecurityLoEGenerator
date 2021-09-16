"""loesite URL Configuration

The `urlpatterns` list routes URLs to views. For more information please see:
    https://docs.djangoproject.com/en/3.2/topics/http/urls/
Examples:
Function views
    1. Add an import:  from my_app import views
    2. Add a URL to urlpatterns:  path('', views.home, name='home')
Class-based views
    1. Add an import:  from other_app.views import Home
    2. Add a URL to urlpatterns:  path('', Home.as_view(), name='home')
Including another URLconf
    1. Import the include() function: from django.urls import include, path
    2. Add a URL to urlpatterns:  path('blog/', include('blog.urls'))
"""
from django.contrib import admin
from django.urls import path
from . import views

urlpatterns = [
    path('admin/', admin.site.urls),
    path('', views.portal_display, name="display"),
    path('ise_LoE', views.portal_display, name="ise_display"),
    path('firepower_LoE', views.fp_display, name="fp_display"),
    path('ise_output_LoE/', views.ise_form_process, name="downloadpage"),
    path('firepower_output_LoE/', views.firepower_form_process, name="downloadpage"),
    path('output_LoE/download', views.file_download, name="download")
]
