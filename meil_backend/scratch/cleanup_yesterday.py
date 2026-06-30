import os
import django
from django.db import transaction
import sys

# Setup Django environment
os.environ.setdefault('DJANGO_SETTINGS_MODULE', 'core.settings')
django.setup()

from matg_attributes.models import MatgAttributeItem

def purge_matg_attributes():
    print(f"--- DATABASE PURGE: ALL MATG ATTRIBUTES ---")
    
    # Identify all records
    matg_items = MatgAttributeItem.objects.all()
    count_matg = matg_items.count()
    
    print(f"Found {count_matg} total MatgAttributeItem records.")
    
    if count_matg == 0:
        print("No records found to delete. Exiting.")
        return

    # List samples
    print("\nSample MatgAttributeItem records to be deleted:")
    for item in matg_items[:10]:
        print(f"  - [ID: {item.id}] {item.attribute_name} (Group: {item.mgrp_code_id})")
            
    # Confirmation
    print("\n" + "="*60)
    print("CRITICAL WARNING: This will permanently DELETE ALL MatgAttributeItem records.")
    print("This action cannot be undone.")
    print("="*60)
    
    try:
        val = input("Confirm FULL PURGE by typing 'yes': ")
        if val.lower() != 'yes':
            print("Deletion cancelled by user.")
            return
    except EOFError:
        print("Non-interactive mode detected. If you want to FORCE deletion, edit the script to bypass this check.")
        return

    with transaction.atomic():
        deleted_count = matg_items.delete()
        print(f"\nPURGE COMPLETE:")
        print(f"  - MatgAttributeItem: {deleted_count[0]} records deleted.")

if __name__ == "__main__":
    purge_matg_attributes()
