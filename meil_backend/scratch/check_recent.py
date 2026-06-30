import os
import sys
import django
from django.utils import timezone
from django.apps import apps
from datetime import timedelta

# Setup Django environment
sys.path.append(os.getcwd())
os.environ.setdefault('DJANGO_SETTINGS_MODULE', 'core.settings')
django.setup()

def check_recent_uploads():
    now = timezone.now()
    yesterday = (now - timedelta(days=1)).date()
    today = now.date()
    
    print(f"Checking for uploads on {yesterday} and {today}...")
    
    found = False
    for model in apps.get_models():
        try:
            time_fields = [f.name for f in model._meta.fields if f.name in ['created', 'created_at', 'date_created']]
            if not time_fields: continue
            
            for field in time_fields:
                for target_date in [yesterday, today]:
                    filter_kwargs = {f"{field}__date": target_date}
                    count = model.objects.filter(**filter_kwargs).count()
                    if count > 0:
                        print(f"Model {model.__name__} has {count} records on {target_date} (field: {field})")
                        found = True
                        items = model.objects.filter(**filter_kwargs)[:5]
                        for item in items:
                            print(f"  - {item}")
        except: pass
            
    if not found:
        print("No records found in the last 2 days.")

if __name__ == "__main__":
    check_recent_uploads()
