import os
import django
from datetime import date, timedelta
from django.apps import apps
from django.utils import timezone

os.environ.setdefault('DJANGO_SETTINGS_MODULE', 'core.settings')
django.setup()

def find_recent_records():
    target_dates = [date(2026, 4, 18), date(2026, 4, 17), date(2026, 4, 16)]
    print(f"Searching for records on dates: {target_dates}")
    
    for model in apps.get_models():
        # Check for different timestamp fields
        timestamp_fields = ['created', 'created_at', 'updated', 'updated_at', 'date_joined']
        fields = [f.name for f in model._meta.fields]
        
        relevant_fields = [f for f in timestamp_fields if f in fields]
        
        if not relevant_fields:
            continue
            
        for t_field in relevant_fields:
            for t_date in target_dates:
                try:
                    query = {f"{t_field}__date": t_date}
                    count = model.objects.filter(**query).count()
                    if count > 0:
                        print(f"MODEL: {model._meta.label} | FIELD: {t_field} | DATE: {t_date} | COUNT: {count}")
                except Exception as e:
                    pass

if __name__ == "__main__":
    find_recent_records()
