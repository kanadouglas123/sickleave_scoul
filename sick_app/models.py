# sick_app/models.py
from django.db import models
from django.contrib.auth.models import User

class Employee(models.Model):
    employee_code = models.CharField(
        max_length=50, unique=True,
        help_text="Unique identifier for the employee (e.g., EMP001)."
    )
    employee_name = models.CharField(
        max_length=100,
        help_text="Full name of the employee."
    )
    department = models.CharField(
        max_length=100,
        help_text="Department where the employee works."
    )
    designation = models.CharField(
        max_length=100,
        help_text="Job title or role of the employee."
    )
    current_total_days = models.IntegerField(
        default=30,
        help_text="Current total available sick leave days."
    )
    additional_sick_leave_days = models.IntegerField(
        default=0,
        help_text="Additional sick leave days granted."
    )
    reason = models.CharField(
        max_length=100,
        help_text="Reason for additional days",
        blank=True, null=True
    )

    def __str__(self):
        return f"{self.employee_name} ({self.employee_code})"

    class Meta:
        ordering = ['employee_name']

class Doctor(models.Model):
    name = models.CharField(max_length=100)
    def __str__(self):
        return self.name

    class Meta:
        db_table = 'doctor'

class SickLeave(models.Model):
    employee = models.ForeignKey(
        Employee, on_delete=models.CASCADE,
        related_name='sick_leaves'
    )
    sick_leave_days = models.IntegerField(default=30)
    days_required = models.IntegerField()
    start_date = models.DateField(null=True, blank=True)
    end_date = models.DateField(null=True, blank=True)
    patient_service = models.CharField(max_length=100)
    gender = models.CharField(max_length=10)
    doctor_remarks = models.TextField()
    recommendation = models.TextField()
    document = models.FileField(upload_to='sick_leave_documents/', null=True, blank=True)
    approved_by = models.ForeignKey(
        Doctor, 
        on_delete=models.SET_NULL, 
        null=True, 
        blank=True,
        help_text="Doctor who approved the sick leave"
    )
    balance_days = models.IntegerField(
        help_text="Remaining balance after deducting required days",
        null=True, blank=True
    )
    created_by = models.ForeignKey(User, on_delete=models.CASCADE)
    created_at = models.DateTimeField(auto_now_add=True)

    def save(self, *args, **kwargs):
        # Compute new balance
        new_balance = (
            self.employee.current_total_days
            - self.days_required
            + self.employee.additional_sick_leave_days
        )

        # Prevent negative balance
        if new_balance < 0:
            new_balance = 0

        # Set balance and update employee
        self.balance_days = new_balance
        self.employee.current_total_days = new_balance
        self.employee.save()

        super(SickLeave, self).save(*args, **kwargs)

    def __str__(self):
        return f"Sick Leave for {self.employee.employee_name} - {self.created_at.strftime('%Y-%m-%d')}"

    class Meta:
        indexes = [
            models.Index(fields=['employee']),
            models.Index(fields=['created_by']),
            models.Index(fields=['created_at'])
        ]
        ordering = ['-created_at']
        verbose_name = 'Sick Leave'
        verbose_name_plural = 'Sick Leaves'