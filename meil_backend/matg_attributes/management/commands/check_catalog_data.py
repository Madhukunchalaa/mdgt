"""
Read-only report to run BEFORE going live:

  * material groups that have two characteristics with the same Sequence
  * characteristic values that still have no CMT Code

    python manage.py check_catalog_data
"""
from django.core.management.base import BaseCommand

from matg_attributes import catalog
from matg_attributes.models import MatgAttributeItem


class Command(BaseCommand):
    help = "Report duplicate sequences and missing CMT codes (changes nothing)."

    def handle(self, *args, **opts):
        dups = catalog.duplicate_sequences()
        self.stdout.write(self.style.MIGRATE_HEADING(f"Duplicate sequences: {len(dups)}"))
        for d in dups:
            self.stdout.write(f"  group {d['mgrp_code']}  sequence {d['sequence']}  ->  {', '.join(d['attributes'])}")

        missing = 0
        self.stdout.write(self.style.MIGRATE_HEADING("Values without a valid CMT Code"))
        for a in MatgAttributeItem.objects.filter(is_deleted=False).order_by("mgrp_code_id", "print_priority"):
            _, errs = catalog.validate_cmt_codes(a.possible_values, a.cmt_codes)
            if errs:
                missing += 1
                self.stdout.write(f"  group {a.mgrp_code_id}  {a.attribute_name}: {len(errs)} problem(s), e.g. {errs[0]}")
        self.stdout.write("")
        if dups or missing:
            self.stdout.write(self.style.WARNING(
                f"{len(dups)} duplicate-sequence case(s), {missing} characteristic(s) need CMT codes."))
        else:
            self.stdout.write(self.style.SUCCESS("Data is clean."))
