from django.contrib import admin
from .models import Employee, SickLeave
from django.utils.html import format_html
import os
from .models import SickLeave, Employee,Doctor


@admin.register(Doctor)
class DoctorAdmin(admin.ModelAdmin):
    list_display = ['name', 'id']  # Display name and ID in the admin list view
    search_fields = ['name']  # Enable search by doctor name
    list_filter = ['name']  # Add filter by name
    ordering = ['name']
    
@admin.register(Employee)
class EmployeeAdmin(admin.ModelAdmin):
    list_display = (
        'employee_code', 'employee_name', 'department',
        'designation', 'current_total_days',
        'additional_sick_leave_days', 'reason'
    )
    search_fields = ('employee_code', 'employee_name', 'department', 'designation')
    list_filter = ('department', 'designation')
    readonly_fields = ('current_total_days',)

    def save_model(self, request, obj, form, change):
        # Update current_total_days with additional days
        obj.current_total_days += obj.additional_sick_leave_days
        if obj.current_total_days < 0:
            obj.current_total_days = 0
        super().save_model(request, obj, form, change)


@admin.register(SickLeave)
class SickLeaveAdmin(admin.ModelAdmin):
    list_display = (
        'get_employee_code', 'get_employee_name', 'get_department', 'get_designation',
        'days_required', 'balance_days', 'start_date', 'end_date',
        'patient_service', 'gender', 'approved_by',
        'document_display',  # <-- added here
        'created_by', 'created_at'
    )
    list_filter = (
        'created_at', 'created_by',
        'employee__department', 'employee__designation',
        'gender', 'patient_service'
    )
    search_fields = (
        'employee__employee_code', 'employee__employee_name',
        'patient_service', 'doctor_remarks', 'approved_by'
    )
    ordering = ('-created_at',)
    raw_id_fields = ('employee', 'created_by',)
    readonly_fields = ('created_by', 'balance_days')

    # Store request for use in document_display
    def get_queryset(self, request):
        self._request = request
        qs = super().get_queryset(request)
        if request.user.is_superuser:
            return qs
        return qs.filter(created_by=request.user)

    # Show document link or preview based on file type
    def document_display(self, obj):
        request = getattr(self, '_request', None)
        user = request.user if request else None

        if not user or (not user.is_superuser and obj.created_by != user):
            return "🔒 Not allowed"

        if obj.document:
            ext = os.path.splitext(obj.document.name)[1].lower()
            if ext in ['.jpg', '.jpeg', '.png', '.gif']:
                return format_html('<a href="{}" target="_blank">📷 View Image</a>', obj.document.url)
            elif ext == '.pdf':
                return format_html('<a href="{}" target="_blank">📄 View PDF</a>', obj.document.url)
            else:
                return format_html('<a href="{}" target="_blank">📁 View Document</a>', obj.document.url)
        return "-"
    document_display.short_description = "Document"

    # Custom column getters
    def get_employee_code(self, obj):
        return obj.employee.employee_code
    get_employee_code.short_description = 'Employee Code'

    def get_employee_name(self, obj):
        return obj.employee.employee_name
    get_employee_name.short_description = 'Employee Name'

    def get_department(self, obj):
        return obj.employee.department
    get_department.short_description = 'Department'

    def get_designation(self, obj):
        return obj.employee.designation
    get_designation.short_description = 'Designation'

