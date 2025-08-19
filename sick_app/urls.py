from django.contrib import admin
from django.urls import path
from sick_app import views
from django.conf import settings
from django.conf.urls.static import static

#
urlpatterns = [

    path('', views.login,name="login"),
    path('base/', views.base, name="base"),
    path('master/', views.master, name="master"),
    path('staff/', views.staff, name="staff"),
    path('logout/', views.logout, name="logout"),
    path('developer/', views.developer, name="developer"),
    path('list/', views.staff_list, name="staff_list"),
    path('fetch-sickleave/', views.fetch_sickleave_by_code, name='fetch_sickleave_by_code'),
    path('masterleave/', views.masterleave, name="masterleave"),
    path('master-submit/', views.master_submit, name="master_submit"),  # For form submission
    # path('load/', views.load, name="load_employee_data"),
    path('upload/', views.upload_excel, name="upload_excel"),
    path('search_employee/', views.search_employee, name='search_employee'),
    path('fetch-employee/', views.fetch_employee, name='fetch_employee'),
    path('report/',views.Report, name='report'),
    path('export-pdf/', views.export_report_pdf, name='export_report_pdf'),    
    path('report/pdf/', views.export_report_pdf, name='report_pdf'),
    path('report/leave/<int:leave_pk>/pdf/', views.export_single_leave_pdf_view, name='export_single_leave_pdf'),
     
]
urlpatterns += static(settings.MEDIA_URL, document_root=settings.MEDIA_ROOT)
