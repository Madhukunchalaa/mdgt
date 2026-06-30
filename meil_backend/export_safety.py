import os
import django
import json
from django.core.serializers.json import DjangoJSONEncoder

# Setup Django environment
os.environ.setdefault('DJANGO_SETTINGS_MODULE', 'core.settings')
django.setup()

from matgroups.models import MatGroup
from itemmaster.models import ItemMaster
from matg_attributes.models import MatgAttributeItem
from supergroups.models import SuperGroup
from MaterialType.models import MaterialType

def export_safety_data():
    print("Starting export of Safety-related materials...")
    
    # 1. Identify Safety MatGroups
    # We search for 'safety' or 'safty' in short and long names
    safety_groups = MatGroup.objects.filter(
        models.Q(mgrp_shortname__icontains='safety') | 
        models.Q(mgrp_longname__icontains='safety') |
        models.Q(mgrp_shortname__icontains='safty') |
        models.Q(mgrp_longname__icontains='safty'),
        is_deleted=False
    )
    
    # Also include items that have 'safety' in their name but might not be in a safety group
    safety_items = ItemMaster.objects.filter(
        models.Q(short_name__icontains='safety') | 
        models.Q(long_name__icontains='safety'),
        is_deleted=False
    )
    
    # Collect all groups from items too
    group_codes = set(safety_groups.values_list('mgrp_code', flat=True))
    group_codes.update(safety_items.values_list('mgrp_code_id', flat=True))
    
    all_safety_groups = MatGroup.objects.filter(mgrp_code__in=group_codes)
    all_safety_items = ItemMaster.objects.filter(mgrp_code__in=group_codes, is_deleted=False)
    
    # 2. Collect related definitions
    sgrp_codes = set(all_safety_groups.values_list('sgrp_code_id', flat=True))
    mat_type_codes = set(all_safety_items.values_list('mat_type_code_id', flat=True))
    
    super_groups = SuperGroup.objects.filter(sgrp_code__in=sgrp_codes)
    material_types = MaterialType.objects.filter(mat_type_code__in=mat_type_codes)
    
    # Attribute definitions
    attribute_items = MatgAttributeItem.objects.filter(mgrp_code__in=group_codes, is_deleted=False)
    
    # 3. Serialize Data
    def serialize_qs(qs):
        data = []
        for obj in qs:
            fields = {}
            for field in obj._meta.fields:
                val = getattr(obj, field.name)
                # Strip user audit fields to be handled on server, but KEEP timestamps
                if field.name in ['createdby', 'updatedby']:
                    continue
                
                # Handle ForeignKeys (save the pk value)
                if isinstance(field, django.db.models.ForeignKey):
                    fields[field.name] = val.pk if val else None
                else:
                    fields[field.name] = val
            data.append({
                'model': f"{obj._meta.app_label}.{obj._meta.model_name}",
                'pk': obj.pk,
                'fields': fields
            })
        return data

    export_data = {
        'supergroups': serialize_qs(super_groups),
        'materialtypes': serialize_qs(material_types),
        'matgroups': serialize_qs(all_safety_groups),
        'attributes': serialize_qs(attribute_items),
        'items': serialize_qs(all_safety_items),
    }

    with open('safety_data.json', 'w') as f:
        json.dump(export_data, f, cls=DjangoJSONEncoder, indent=2)

    print(f"Export complete. Total items exported: {all_safety_items.count()}")
    print(f"Total groups exported: {all_safety_groups.count()}")
    print("Data saved to safety_data.json")

if __name__ == "__main__":
    from django.db import models
    export_safety_data()
