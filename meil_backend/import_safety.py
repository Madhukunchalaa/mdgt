import os
import django
import json
import sys

# Setup Django environment
os.environ.setdefault('DJANGO_SETTINGS_MODULE', 'core.settings')
django.setup()

from django.apps import apps
from Employee.models import Employee

def import_safety_data(json_file):
    if not os.path.exists(json_file):
        print(f"ERROR: {json_file} not found.")
        return

    print(f"Starting import from {json_file}...")
    
    # 1. Get Shanti's User
    try:
        shanti = Employee.objects.get(email='shanti@meil.com')
        print(f"Found Shanti (emp_id: {shanti.emp_id}). All records will be attributed to her.")
    except Employee.DoesNotExist:
        print("ERROR: User 'shanti@meil.com' not found on this server. Please create her first.")
        return

    with open(json_file, 'r') as f:
        data = json.load(f)

    def import_model_data(model_name, items):
        Model = apps.get_model(model_name)
        count = 0
        for item in items:
            pk = item['pk']
            fields = item['fields']
            
            # Map ForeignKeys if needed
            # (In this schema, we mostly match by code which is the PK)
            
            # Special handling for Audit fields
            fields['createdby'] = shanti
            fields['updatedby'] = shanti
            
            # Use update_or_create to avoid duplicates
            pk_name = Model._meta.pk.name
            
            if pk_name in fields:
                fields.pop(pk_name)

            # We use an optimized approach to ensure 'created' isn't null
            obj = Model.objects.filter(**{pk_name: pk}).first()
            if obj:
                # Update existing and ensure not deleted
                for key, value in fields.items():
                    setattr(obj, key, value)
                if hasattr(obj, 'is_deleted'):
                    obj.is_deleted = False
                obj.save()
            else:
                # Create new
                obj = Model(**{pk_name: pk})
                for key, value in fields.items():
                    setattr(obj, key, value)
                
                if hasattr(obj, 'is_deleted'):
                    obj.is_deleted = False
                
                # If 'created' is still missing from fields (because it wasn't in JSON)
                # set it to now
                if not getattr(obj, 'created', None):
                    from django.utils import timezone
                    obj.created = timezone.now()
                
                obj.save()
            count += 1
        print(f"Processed {count} records in {model_name}")

    # Load in order of dependencies
    # SuperGroups -> MaterialTypes -> MatGroups -> Attributes -> ItemMaster
    print("\n--- Phase 1: SuperGroups ---")
    import_model_data('supergroups.supergroup', data['supergroups'])
    
    print("\n--- Phase 2: MaterialTypes ---")
    import_model_data('MaterialType.materialtype', data['materialtypes'])
    
    print("\n--- Phase 3: MatGroups ---")
    import_model_data('matgroups.matgroup', data['matgroups'])
    
    print("\n--- Phase 4: Attribute Definitions ---")
    import_model_data('matg_attributes.matgattributeitem', data['attributes'])
    
    print("\n--- Phase 5: Item Master (Materials) ---")
    import_model_data('itemmaster.itemmaster', data['items'])

    print("\nMigration Complete!")

if __name__ == "__main__":
    import_safety_data('safety_data.json')
