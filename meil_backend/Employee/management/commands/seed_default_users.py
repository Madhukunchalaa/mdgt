"""
Management command to seed a fresh database with default roles, company,
email domain, and one user per role.

Usage:
    python manage.py seed_default_users

Idempotent — safe to run multiple times; existing records are skipped.
"""

from django.core.management.base import BaseCommand
from django.contrib.auth.hashers import make_password


DEFAULT_USERS = [
    {
        "emp_name": "Super Admin",
        "email": "superadmin@meil.com",
        "password": "SuperAdmin@123",
        "role": "SUPERADMIN",
        "designation": "Super Administrator",
        "phone_number": "9000000001",
    },
    {
        "emp_name": "Admin User",
        "email": "admin@meil.com",
        "password": "Admin@123",
        "role": "ADMIN",
        "designation": "Administrator",
        "phone_number": "9000000002",
    },
    {
        "emp_name": "MDGT Operator",
        "email": "mdgt@meil.com",
        "password": "Mdgt@123",
        "role": "MDGT",
        "designation": "MDGT Operator",
        "phone_number": "9000000003",
    },
    {
        "emp_name": "Regular User",
        "email": "user@meil.com",
        "password": "User@123",
        "role": "USER",
        "designation": "User",
        "phone_number": "9000000004",
    },
]

ROLES = [
    {"role_name": "SUPERADMIN", "role_priority": 1, "can_create": True,  "can_update": True,  "can_delete": True,  "can_export": True},
    {"role_name": "ADMIN",      "role_priority": 2, "can_create": True,  "can_update": True,  "can_delete": True,  "can_export": True},
    {"role_name": "MDGT",       "role_priority": 3, "can_create": True,  "can_update": True,  "can_delete": False, "can_export": True},
    {"role_name": "USER",       "role_priority": 4, "can_create": False, "can_update": False, "can_delete": False, "can_export": True},
]


class Command(BaseCommand):
    help = "Seed the database with default company, roles, and one user per role"

    def handle(self, *args, **options):
        from Company.models import Company
        from EmailDomain.models import EmailDomain
        from Users.models import UserRole
        from Employee.models import Employee

        self.stdout.write(self.style.MIGRATE_HEADING("\n=== Seeding default data ===\n"))

        # 1. Default company
        company, created = Company.objects.get_or_create(
            company_name="MEIL",
            defaults={"contact": "9000000000", "is_deleted": False},
        )
        self._log("Company", "MEIL", created)

        # 2. Email domain (required for user login via the register flow)
        domain, created = EmailDomain.objects.get_or_create(
            domain_name="meil.com",
            defaults={"is_deleted": False},
        )
        self._log("EmailDomain", "meil.com", created)

        # 3. Roles
        roles = {}
        for rd in ROLES:
            role, created = UserRole.objects.get_or_create(
                role_name=rd["role_name"],
                defaults={
                    "role_priority": rd["role_priority"],
                    "can_create": rd["can_create"],
                    "can_update": rd["can_update"],
                    "can_delete": rd["can_delete"],
                    "can_export": rd["can_export"],
                    "is_deleted": False,
                },
            )
            roles[rd["role_name"]] = role
            self._log("Role", rd["role_name"], created)

        # 4. Employees — one per role
        for ud in DEFAULT_USERS:
            emp, created = Employee.objects.get_or_create(
                email=ud["email"],
                is_deleted=False,
                defaults={
                    "emp_name": ud["emp_name"],
                    "password": make_password(ud["password"]),
                    "role": roles[ud["role"]],
                    "company_name": company,
                    "designation": ud["designation"],
                    "phone_number": ud["phone_number"],
                    "is_email_verified": True,
                    "is_deleted": False,
                },
            )
            if created:
                # Self-reference for audit trail
                emp.createdby = emp
                emp.updatedby = emp
                emp.save()
            self._log("Employee", f"{ud['email']}  [{ud['role']}]", created)

        self.stdout.write(self.style.SUCCESS("\n=== Seed complete! ===\n"))
        self.stdout.write("Default login credentials:")
        self.stdout.write("  SUPERADMIN : superadmin@meil.com  /  SuperAdmin@123")
        self.stdout.write("  ADMIN      : admin@meil.com       /  Admin@123")
        self.stdout.write("  MDGT       : mdgt@meil.com        /  Mdgt@123")
        self.stdout.write("  USER       : user@meil.com        /  User@123")
        self.stdout.write("")

    def _log(self, kind, name, created):
        if created:
            self.stdout.write(self.style.SUCCESS(f"  [CREATED] {kind}: {name}"))
        else:
            self.stdout.write(f"  [EXISTS]  {kind}: {name}")
