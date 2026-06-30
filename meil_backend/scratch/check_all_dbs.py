import os
import sys
import django
from django.utils import timezone
from django.apps import apps
from django.conf import settings

# Override database settings to use sqlite3
os.environ.setdefault('DJANGO_SETTINGS_MODULE', 'core.settings')

# We need to configure settings before django.setup() if we want to override them
from django.conf import settings
if not settings.configured:
    # This might not work if core.settings is already loaded
    pass

# Better approach: modify the settings after setup or before
def check_database(db_name, db_engine, db_file=None):
    print(f"\n--- Checking {db_name} ({db_engine}) ---")
    
    # Configure database
    new_db_settings = {
        'default': {
            'ENGINE': db_engine,
        }
    }
    if db_file:
        new_db_settings['default']['NAME'] = db_file
    else:
        # For postgres, it will use whatever is in env or default
        import dj_database_url
        new_db_settings['default'] = dj_database_url.config(
            default=os.environ.get('DATABASE_URL', 'postgresql://postgres:123@localhost:5432/mdgt')
        )

    # Apply settings
    from django.db import connections
    connections.databases['default'] = new_db_settings['default']
    
    today = timezone.now().date()
    found = False
    for model in apps.get_models():
        try:
            time_fields = [f.name for f in model._meta.fields if f.name in ['created', 'created_at', 'date_created']]
            if not time_fields: continue
            for field in time_fields:
                filter_kwargs = {f"{field}__date": today}
                count = model.objects.filter(**filter_kwargs).count()
                if count > 0:
                    print(f"Model {model.__name__} has {count} records created today (field: {field})")
                    found = True
        except: pass
    if not found: print("No records found created today.")

if __name__ == "__main__":
    # Setup standard django
    sys.path.append(os.getcwd())
    django.setup()
    
    # Check Postgres (default)
    check_database("PostgreSQL", "django.db.backends.postgresql")
    
    # Check SQLite
    sqlite_file = os.path.join(os.getcwd(), 'db.sqlite3')
    if os.path.exists(sqlite_file):
        check_database("SQLite", "django.db.backends.sqlite3", sqlite_file)
    else:
        print("\ndb.sqlite3 not found.")
