import os
import sys
import django
from django.utils import timezone
from django.apps import apps

# Add the current directory to sys.path so it can find the 'core' module
sys.path.append(os.getcwd())

# Setup Django environment
os.environ.setdefault('DJANGO_SETTINGS_MODULE', 'core.settings')
django.setup()

def check_today_uploads():
    today = timezone.now().date()
    print(f"Checking for uploads on {today}...")
    
    found = False
    for model in apps.get_models():
        try:
            # Check for fields that represent creation time
            time_fields = [f.name for f in model._meta.fields if f.name in ['created', 'created_at', 'date_created', 'timestamp']]
            if not time_fields:
                continue
                
            for field in time_fields:
                filter_kwargs = {f"{field}__date": today}
                count = model.objects.filter(**filter_kwargs).count()
                if count > 0:
                    print(f"Model {model.__name__} has {count} records created today (field: {field})")
                    found = True
                    # Print first 5 items
                    items = model.objects.filter(**filter_kwargs)[:5]
                    for item in items:
                        print(f"  - {item}")
        except Exception as e:
            # print(f"Error checking model {model.__name__}: {e}")
            pass
            
    if not found:
        print("No records found created today in any model with 'created' or similar fields.")

if __name__ == "__main__":
    check_today_uploads()
