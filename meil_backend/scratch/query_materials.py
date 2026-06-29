import os
import sys
import django

# Add meil_backend directory to PYTHONPATH
sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

os.environ.setdefault('DJANGO_SETTINGS_MODULE', 'core.settings')
django.setup()

from itemmaster.models import ItemMaster

items = ItemMaster.objects.filter(is_deleted=False)[:5]
print(f"Total active items: {ItemMaster.objects.filter(is_deleted=False).count()}")
for item in items:
    print(f"Local ID: {item.local_item_id} | SAP ID: {item.sap_item_id} | Short: {item.short_name} | Long: {item.long_name} | Search: {item.search_text}")
