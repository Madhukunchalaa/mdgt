"""
Fill / refresh the Catalog Number of existing materials, and put their
attributes in sequence order.

    python manage.py backfill_catalog_numbers            # all groups
    python manage.py backfill_catalog_numbers --group 300100101
    python manage.py backfill_catalog_numbers --dry-run
"""
from django.core.management.base import BaseCommand

from itemmaster.models import ItemMaster
from matg_attributes import catalog
from matgroups.models import MatGroup


class Command(BaseCommand):
    help = "Recalculate catalog_number for existing materials."

    def add_arguments(self, parser):
        parser.add_argument("--group", help="Only this material group code")
        parser.add_argument("--dry-run", action="store_true", help="Show what would change, save nothing")

    def handle(self, *args, **opts):
        groups = MatGroup.objects.all()
        if opts["group"]:
            groups = groups.filter(mgrp_code=opts["group"])
        changed = skipped = with_warnings = 0
        for group in groups:
            for item in ItemMaster.objects.filter(mgrp_code=group, is_deleted=False).iterator():
                info = catalog.catalog_breakdown(group, item.attributes)
                if info["warnings"]:
                    with_warnings += 1
                if info["catalog_number"] == item.catalog_number:
                    skipped += 1
                    continue
                changed += 1
                if not opts["dry_run"]:
                    ItemMaster.objects.filter(pk=item.pk).update(catalog_number=info["catalog_number"])
        verb = "would update" if opts["dry_run"] else "updated"
        self.stdout.write(self.style.SUCCESS(
            f"{verb} {changed} item(s); {skipped} already correct; "
            f"{with_warnings} item(s) are incomplete (a value has no CMT code, so no number was generated)."))
