from matg_attributes.models import MatgAttributeItem
from itemmaster.models import ItemMaster
from matgroups.models import MatGroup
from datetime import date, timedelta
import django
import os

os.environ.setdefault('DJANGO_SETTINGS_MODULE', 'core.settings')
django.setup()

today = date(2026, 4, 18)
print(f"Checking records around {today}")
for i in range(7):
    d = today - timedelta(days=i)
    c_matg = MatgAttributeItem.objects.filter(created__date=d).count()
    c_item = ItemMaster.objects.filter(created__date=d).count()
    c_group = MatGroup.objects.filter(created__date=d).count()
    print(f"{d}: MatgAttributeItem={c_matg}, ItemMaster={c_item}, MatGroup={c_group}")
