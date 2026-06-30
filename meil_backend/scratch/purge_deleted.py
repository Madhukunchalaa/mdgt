import os
import django

# Setup Django environment
os.environ.setdefault('DJANGO_SETTINGS_MODULE', 'core.settings')
django.setup()

from matg_attributes.models import MatgAttributeItem
from matgroups.models import MatGroup

def purge_deleted_attributes():
    deleted_attrs = MatgAttributeItem.objects.filter(is_deleted=True)
    count = deleted_attrs.count()
    if count > 0:
        deleted_attrs.delete()
        print(f"SUCCESS: Permanently purged {count} soft-deleted attribute(s) from the database.")
    else:
        print("No soft-deleted attributes found to purge.")

if __name__ == "__main__":
    purge_deleted_attributes()
