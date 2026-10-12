"""
Tests for: CMT Code, unique Sequence, ordered description, Catalog Number.

The expected Catalog Numbers come straight from the client's sample sheet
("Sample Material group code_piping", Piping tab, FLANGE group 300100101).
Run:  python manage.py test matg_attributes
"""
import datetime
import json

import jwt
from django.conf import settings
from django.core.management import call_command
from django.test import TestCase, Client

from Employee.models import Employee
from MaterialType.models import MaterialType
from matgroups.models import MatGroup
from matg_attributes.models import MatgAttributeItem
from matg_attributes import catalog
from itemmaster.models import ItemMaster

GROUP = "300100101"

# Characteristic table of the FLANGE group, exactly as in the sample sheet.
# (name, sequence, values, codes).  Schedule has no values for flanges -> code 00.
FLANGE_CHARACTERISTICS = [
    ("PART TYPE", 10, ["FLANGE"], {"FLANGE": "01"}),
    ("SUB TYPE", 20, ["BLIND", "FLANGE", "WNRF"], {"BLIND": "01", "FLANGE": "02", "WNRF": "03"}),
    ("SIZE", 30, ["0.5", "0.75", "1", "1.5", "2", "2.5", "3", "4"],
        {"0.5": "01", "0.75": "02", "1": "03", "1.5": "04", "2": "05", "2.5": "06", "3": "07", "4": "08"}),
    ("SCHEDULE", 40, [], {}),
    ("RATING", 50, ["150", "300", "600"], {"150": "01", "300": "02", "600": "03"}),
    ("MOC", 60, ["SS", "AS", "CS"], {"SS": "01", "AS": "02", "CS": "03"}),
    ("STANDARD", 70, ["B16.5", "B16.48"], {"B16.5": "01", "B16.48": "02"}),
]

# Rows 4-12 of the Piping tab: (selected values, catalog number in the sheet)
SHEET_ROWS = [
    (("FLANGE", "BLIND", "0.75", None, "150", "SS", "B16.5"), "01010200010101"),
    (("FLANGE", "BLIND", "1", None, "150", "SS", "B16.5"), "01010300010101"),
    (("FLANGE", "BLIND", "1.5", None, "150", "SS", "B16.5"), "01010400010101"),
    (("FLANGE", "BLIND", "2", None, "300", "AS", "B16.48"), "01010500020202"),
    (("FLANGE", "BLIND", "2.5", None, "300", "AS", "B16.48"), "01010600020202"),
    (("FLANGE", "BLIND", "3", None, "300", "AS", "B16.48"), "01010700020202"),
    (("FLANGE", "BLIND", "4", None, "600", "CS", "B16.48"), "01010800030302"),
    (("FLANGE", "FLANGE", "3", None, "300", "AS", "B16.48"), "01020700020202"),
    (("FLANGE", "WNRF", "4", None, "600", "CS", "B16.48"), "01030800030302"),
]


class Base(TestCase):
    @classmethod
    def setUpTestData(cls):
        call_command("seed_default_users", verbosity=0)
        cls.emp = Employee.objects.get(email="admin@meil.com")
        MaterialType.objects.create(mat_type_code="ZCON", mat_type_desc="Construction")
        cls.group = MatGroup.objects.create(mgrp_code=GROUP, mgrp_shortname="FLANGE", mgrp_longname="Con Mechanical Piping FLANGE")
        cls.other_group = MatGroup.objects.create(mgrp_code="300100102", mgrp_shortname="GASKET", mgrp_longname="Gasket")

    def setUp(self):
        token = jwt.encode(
            {"emp_id": self.emp.emp_id, "email": self.emp.email, "role": "ADMIN",
             "exp": datetime.datetime.utcnow() + datetime.timedelta(hours=1)},
            settings.SECRET_KEY, algorithm="HS256")
        self.c = Client(HTTP_AUTHORIZATION=f"Bearer {token}")

    # -- helpers --------------------------------------------------------
    def post(self, url, payload):
        return self.c.post(url, json.dumps(payload), content_type="application/json")

    def put(self, url, payload):
        return self.c.put(url, json.dumps(payload), content_type="application/json")

    def add_char(self, name, seq, values, codes, group=GROUP):
        return self.post("/matgattribute/create/", {
            "mgrp_code": group,
            "attributes": [{"attribute_name": name, "sequence": seq, "possible_values": values, "cmt_codes": codes}],
        })

    def load_flange_table(self):
        for name, seq, values, codes in FLANGE_CHARACTERISTICS:
            r = self.add_char(name, seq, values, codes)
            self.assertEqual(r.status_code, 201, r.content)

    def make_item(self, selected, scrambled=False):
        names = [c[0] for c in FLANGE_CHARACTERISTICS]
        attrs = {n: v for n, v in zip(names, selected) if v}
        if scrambled:   # send in reverse order on purpose
            attrs = dict(reversed(list(attrs.items())))
        return self.post("/itemmaster/create/", {
            "mat_type_code": "ZCON", "mgrp_code": GROUP, "item_desc": "FLANGE",
            "sap_name": "FLANGE", "attributes": attrs,
        })


# ===========================================================================
# 1. CMT Code
# ===========================================================================
class CmtCodeTests(Base):
    def test_valid_code_is_saved(self):
        r = self.add_char("SIZE", 10, ["0.5", "1"], {"0.5": "01", "1": "02"})
        self.assertEqual(r.status_code, 201)
        self.assertEqual(MatgAttributeItem.objects.get(attribute_name="SIZE").cmt_codes, {"0.5": "01", "1": "02"})

    def test_code_is_mandatory(self):
        r = self.add_char("SIZE", 10, ["0.5", "1"], {})
        self.assertEqual(r.status_code, 400)
        self.assertEqual(r.json()["code"], "INVALID_CMT_CODE")
        self.assertFalse(MatgAttributeItem.objects.exists())

    def test_one_value_missing_a_code(self):
        r = self.add_char("SIZE", 10, ["0.5", "1"], {"0.5": "01"})
        self.assertEqual(r.status_code, 400)
        self.assertIn("'1'", r.json()["error"])

    def test_must_be_exactly_two_digits_numeric(self):
        for bad in ["1", "001", "AB", "1A", "", " ", "1.5", "-1", "০১"]:
            r = self.add_char("SIZE", 10, ["0.5"], {"0.5": bad})
            self.assertEqual(r.status_code, 400, f"{bad!r} should be rejected")
            MatgAttributeItem.objects.all().delete()

    def test_zero_codes_allowed(self):
        self.assertEqual(self.add_char("SCHEDULE", 10, ["NONE"], {"NONE": "00"}).status_code, 201)

    def test_duplicate_code_in_same_characteristic_rejected(self):
        r = self.add_char("SIZE", 10, ["0.5", "1"], {"0.5": "01", "1": "01"})
        self.assertEqual(r.status_code, 400)
        self.assertIn("01", r.json()["error"])

    def test_same_code_allowed_in_different_characteristics(self):
        self.assertEqual(self.add_char("SIZE", 10, ["0.5"], {"0.5": "01"}).status_code, 201)
        self.assertEqual(self.add_char("RATING", 20, ["150"], {"150": "01"}).status_code, 201)

    def test_update_cannot_drop_code_for_new_value(self):
        self.add_char("SIZE", 10, ["0.5"], {"0.5": "01"})
        cid = MatgAttributeItem.objects.get().id
        r = self.put(f"/matgattribute/update/{cid}/", {"possible_values": ["0.5", "1"]})
        self.assertEqual(r.status_code, 400)
        r = self.put(f"/matgattribute/update/{cid}/", {"possible_values": ["0.5", "1"], "cmt_codes": {"0.5": "01", "1": "02"}})
        self.assertEqual(r.status_code, 200)

    def test_bad_request_creates_nothing(self):
        """Second characteristic is invalid -> the first must not be saved either."""
        r = self.post("/matgattribute/create/", {"mgrp_code": GROUP, "attributes": [
            {"attribute_name": "SIZE", "sequence": 10, "possible_values": ["1"], "cmt_codes": {"1": "01"}},
            {"attribute_name": "RATING", "sequence": 20, "possible_values": ["150"], "cmt_codes": {"150": "XX"}},
        ]})
        self.assertEqual(r.status_code, 400)
        self.assertEqual(MatgAttributeItem.objects.count(), 0)

    def test_list_returns_codes(self):
        self.add_char("SIZE", 10, ["0.5"], {"0.5": "01"})
        row = self.c.get("/matgattribute/list/?mgrp_code=" + GROUP).json()[0]
        self.assertEqual(row["cmt_codes"], {"0.5": "01"})


# ===========================================================================
# 2. Unique sequence
# ===========================================================================
class SequenceTests(Base):
    def test_duplicate_on_create_rejected(self):
        self.add_char("SIZE", 10, ["1"], {"1": "01"})
        r = self.add_char("RATING", 10, ["150"], {"150": "01"})
        self.assertEqual(r.status_code, 400)
        self.assertEqual(r.json()["code"], "DUPLICATE_SEQUENCE")
        self.assertIn("SIZE", r.json()["error"])

    def test_duplicate_inside_one_request_rejected(self):
        r = self.post("/matgattribute/create/", {"mgrp_code": GROUP, "attributes": [
            {"attribute_name": "SIZE", "sequence": 5, "possible_values": ["1"], "cmt_codes": {"1": "01"}},
            {"attribute_name": "RATING", "sequence": 5, "possible_values": ["150"], "cmt_codes": {"150": "01"}},
        ]})
        self.assertEqual(r.status_code, 400)
        self.assertEqual(MatgAttributeItem.objects.count(), 0)

    def test_same_sequence_allowed_in_different_groups(self):
        self.assertEqual(self.add_char("SIZE", 10, ["1"], {"1": "01"}).status_code, 201)
        self.assertEqual(self.add_char("SIZE", 10, ["1"], {"1": "01"}, group="300100102").status_code, 201)

    def test_duplicate_on_update_rejected(self):
        self.add_char("SIZE", 10, ["1"], {"1": "01"})
        self.add_char("RATING", 20, ["150"], {"150": "01"})
        rating = MatgAttributeItem.objects.get(attribute_name="RATING")
        r = self.put(f"/matgattribute/update/{rating.id}/", {"sequence": 10})
        self.assertEqual(r.status_code, 400)
        self.assertEqual(r.json()["code"], "DUPLICATE_SEQUENCE")
        rating.refresh_from_db()
        self.assertEqual(rating.print_priority, 20)

    def test_keeping_own_sequence_on_update_is_fine(self):
        self.add_char("SIZE", 10, ["1"], {"1": "01"})
        size = MatgAttributeItem.objects.get()
        self.assertEqual(self.put(f"/matgattribute/update/{size.id}/", {"sequence": 10}).status_code, 200)

    def test_invalid_sequence_values_rejected(self):
        self.add_char("SIZE", 10, ["1"], {"1": "01"})
        size = MatgAttributeItem.objects.get()
        for bad in [0, -3, 1.5, "abc", "2.7"]:
            r = self.put(f"/matgattribute/update/{size.id}/", {"sequence": bad})
            self.assertEqual(r.status_code, 400, f"{bad!r}")

    def test_old_characteristic_without_sequence_can_still_be_edited(self):
        """The edit form sends print_priority 0 for characteristics that never had one."""
        MatgAttributeItem.objects.create(mgrp_code=self.group, attribute_name="OLD", possible_values=["A"], cmt_codes={"A": "01"})
        cid = MatgAttributeItem.objects.get().id
        r = self.put(f"/matgattribute/update/{cid}/", {"attribute_name": "OLD", "possible_values": ["A"],
                                                       "cmt_codes": {"A": "01"}, "print_priority": 0})
        self.assertEqual(r.status_code, 200, r.content)
        self.assertIsNone(MatgAttributeItem.objects.get(id=cid).print_priority)

    def test_cannot_blank_out_an_existing_sequence(self):
        self.add_char("SIZE", 10, ["1"], {"1": "01"})
        cid = MatgAttributeItem.objects.get().id
        self.assertEqual(self.put(f"/matgattribute/update/{cid}/", {"print_priority": 0}).status_code, 400)

    def test_deleted_characteristic_frees_its_sequence(self):
        self.add_char("SIZE", 10, ["1"], {"1": "01"})
        size = MatgAttributeItem.objects.get()
        self.c.delete(f"/matgattribute/delete/{size.id}/")
        self.assertEqual(self.add_char("RATING", 10, ["150"], {"150": "01"}).status_code, 201)

    def test_restore_blocked_when_sequence_taken(self):
        self.add_char("SIZE", 10, ["1"], {"1": "01"})
        size = MatgAttributeItem.objects.get()
        self.c.delete(f"/matgattribute/delete/{size.id}/")
        self.add_char("RATING", 10, ["150"], {"150": "01"})
        r = self.c.post(f"/matgattribute/restore/{size.id}/")
        self.assertEqual(r.status_code, 400)
        self.assertEqual(r.json()["code"], "DUPLICATE_SEQUENCE")

    def test_auto_sequence_when_missing(self):
        self.add_char("SIZE", None, ["1"], {"1": "01"})
        self.add_char("RATING", None, ["150"], {"150": "01"})
        seqs = sorted(MatgAttributeItem.objects.values_list("print_priority", flat=True))
        self.assertEqual(seqs, [10, 20])

    def test_conflict_report_finds_existing_bad_data(self):
        MatgAttributeItem.objects.create(mgrp_code=self.group, attribute_name="A", print_priority=5, possible_values=[])
        MatgAttributeItem.objects.create(mgrp_code=self.group, attribute_name="B", print_priority=5, possible_values=[])
        data = self.c.get("/matgattribute/sequence-conflicts/").json()
        self.assertEqual(data["count"], 1)
        self.assertEqual(sorted(data["conflicts"][0]["attributes"]), ["A", "B"])


# ===========================================================================
# 3 + 4. Description order and Catalog Number  (the client's sample sheet)
# ===========================================================================
class CatalogNumberTests(Base):
    def test_every_row_of_the_sample_sheet(self):
        self.load_flange_table()
        for selected, expected in SHEET_ROWS:
            r = self.make_item(selected)
            self.assertEqual(r.status_code, 201, r.content)
            self.assertEqual(r.json()["catalog_number"], expected, selected)

    def test_order_follows_sequence_even_if_request_is_scrambled(self):
        self.load_flange_table()
        selected, expected = SHEET_ROWS[0]
        r = self.make_item(selected, scrambled=True)
        body = r.json()
        self.assertEqual(body["catalog_number"], expected)
        self.assertEqual(body["short_name"], "FLANGE, FLANGE, BLIND, 0.75, 150, SS, B16.5")
        self.assertEqual(list(catalog.order_attributes(self.group, ItemMaster.objects.get(pk=body["local_item_id"]).attributes)),
                         ["PART TYPE", "SUB TYPE", "SIZE", "RATING", "MOC", "STANDARD"])

    def test_changing_sequence_changes_description_and_catalog_number_together(self):
        self.load_flange_table()
        item_id = self.make_item(SHEET_ROWS[0][0]).json()["local_item_id"]
        size = MatgAttributeItem.objects.get(attribute_name="SIZE")
        part = MatgAttributeItem.objects.get(attribute_name="PART TYPE")
        # swap SIZE (30) and PART TYPE (10): park PART TYPE on 35 first so sequences stay unique
        self.assertEqual(self.put(f"/matgattribute/update/{part.id}/", {"sequence": 35}).status_code, 200)
        item = ItemMaster.objects.get(pk=item_id)
        # SIZE(30) now comes before PART TYPE(35): SUB 01, SIZE 02, PART 01, SCHED 00, RATING 01, MOC 01, STD 01
        self.assertEqual(item.catalog_number, "01" "02" "01" "00" "01" "01" "01")
        self.assertEqual(self.make_item(SHEET_ROWS[1][0]).json()["catalog_number"], "01" "03" "01" "00" "01" "01" "01")

    def test_update_item_rebuilds_catalog_number(self):
        self.load_flange_table()
        item_id = self.make_item(SHEET_ROWS[0][0]).json()["local_item_id"]
        attrs = dict(zip([c[0] for c in FLANGE_CHARACTERISTICS], SHEET_ROWS[3][0]))
        attrs = {k: v for k, v in attrs.items() if v}
        r = self.put(f"/itemmaster/update/{item_id}/", {"attributes": attrs})
        self.assertEqual(r.status_code, 200, r.content)
        self.assertEqual(r.json()["catalog_number"], SHEET_ROWS[3][1])

    def test_editing_a_code_refreshes_existing_items(self):
        self.load_flange_table()
        item_id = self.make_item(SHEET_ROWS[0][0]).json()["local_item_id"]
        moc = MatgAttributeItem.objects.get(attribute_name="MOC")
        r = self.put(f"/matgattribute/update/{moc.id}/", {"cmt_codes": {"SS": "09", "AS": "02", "CS": "03"}})
        self.assertEqual(r.json()["catalog_numbers_refreshed"], 1)
        # SS is now 09:  PART 01, SUB 01, SIZE 02, SCHED 00, RATING 01, MOC 09, STD 01
        self.assertEqual(ItemMaster.objects.get(pk=item_id).catalog_number, "01010200010901")

    def test_blank_characteristic_gets_00(self):
        self.load_flange_table()
        r = self.make_item(("FLANGE", "BLIND", None, None, None, None, None))
        self.assertEqual(r.json()["catalog_number"], "01" "01" "00" "00" "00" "00" "00")

    def test_hidden_characteristics_are_left_out(self):
        self.load_flange_table()
        self.add_char("BU", 80, ["X"], {"X": "05"})
        r = self.make_item(SHEET_ROWS[0][0])
        self.assertEqual(r.json()["catalog_number"], SHEET_ROWS[0][1])   # still 14 digits

    def test_value_without_code_gives_warning_and_no_number(self):
        """A half-built number that looks valid is worse than none."""
        self.load_flange_table()
        info = catalog.catalog_breakdown(self.group, {"PART TYPE": "FLANGE", "SIZE": "99"})
        self.assertIsNone(info["catalog_number"])
        self.assertFalse(info["complete"])
        self.assertTrue(any("99" in w for w in info["warnings"]))

    def test_item_with_uncoded_value_is_saved_without_catalog_number(self):
        self.load_flange_table()
        MatgAttributeItem.objects.filter(attribute_name="MOC").update(cmt_codes={"SS": "01", "AS": "02"})  # CS lost its code
        r = self.make_item(("FLANGE", "BLIND", "1", None, "150", "CS", "B16.5"))
        self.assertEqual(r.status_code, 201)
        self.assertIsNone(r.json()["catalog_number"])
        self.assertTrue(any("CS" in w for w in r.json()["catalog_warnings"]))

    def test_nothing_selected_gives_no_number(self):
        self.load_flange_table()
        self.assertIsNone(catalog.build_catalog_number(self.group, {}))

    def test_numeric_style_values_with_uom_are_matched(self):
        self.add_char("SIZE", 10, ["2"], {"2": "05"})
        MatgAttributeItem.objects.filter(attribute_name="SIZE").update(uom="inch")
        info = catalog.catalog_breakdown(self.group, {"SIZE": {"value": "2", "uom": "inch"}})
        self.assertEqual(info["catalog_number"], "05")
        info = catalog.catalog_breakdown(self.group, {"SIZE": "2 inch"})
        self.assertEqual(info["catalog_number"], "05")

    def test_preview_endpoint(self):
        self.load_flange_table()
        selected, expected = SHEET_ROWS[0]
        attrs = {n: v for n, v in zip([c[0] for c in FLANGE_CHARACTERISTICS], selected) if v}
        data = self.post("/matgattribute/catalog-preview/", {"mgrp_code": GROUP, "attributes": attrs}).json()
        self.assertEqual(data["catalog_number"], expected)
    def test_search_by_catalog_number(self):
        self.load_flange_table()
        self.make_item(SHEET_ROWS[0][0])
        # Search by exact 14-digit catalog number
        r = self.c.get("/itemmaster/list/?search=01010200010101")
        self.assertEqual(r.status_code, 200)
        self.assertEqual(r.json()["count"], 1)
        self.assertEqual(r.json()["results"][0]["catalog_number"], "01010200010101")

        # Search without leading zero (Excel stripped format)
        r2 = self.c.get("/itemmaster/list/?search=1010200010101")
        self.assertEqual(r2.status_code, 200)
        self.assertEqual(r2.json()["count"], 1)
        self.assertEqual(r2.json()["results"][0]["catalog_number"], "01010200010101")


# ===========================================================================
# 5. Bulk upload of characteristics
# ===========================================================================
class UploadTests(Base):
    def run_upload(self, rows):
        from uploads.views import handle_matgattribute_phase_1
        return handle_matgattribute_phase_1(rows, None)

    def row(self, **kw):
        base = {"mgrp_code": GROUP, "attribute_name": "SIZE", "possible_values": "0.5, 1",
                "cmt_codes": "01, 02", "print_priority": 10}
        base.update(kw)
        return base

    def test_positional_codes(self):
        res = self.run_upload([self.row()])
        self.assertEqual(res["errors"], [])
        self.assertEqual(MatgAttributeItem.objects.get().cmt_codes, {"0.5": "01", "1": "02"})

    def test_pair_style_codes(self):
        res = self.run_upload([self.row(cmt_codes="1=02, 0.5=01")])
        self.assertEqual(res["errors"], [])
        self.assertEqual(MatgAttributeItem.objects.get().cmt_codes, {"0.5": "01", "1": "02"})

    def test_excel_numbers_are_padded(self):
        res = self.run_upload([self.row(cmt_codes="1, 2")])        # Excel turned 01 into 1
        self.assertEqual(res["errors"], [])
        self.assertEqual(MatgAttributeItem.objects.get().cmt_codes, {"0.5": "01", "1": "02"})

    def test_missing_codes_rejected(self):
        res = self.run_upload([self.row(cmt_codes=None)])
        self.assertEqual(len(res["errors"]), 1)
        self.assertEqual(MatgAttributeItem.objects.count(), 0)

    def test_wrong_number_of_codes_rejected(self):
        res = self.run_upload([self.row(cmt_codes="01")])
        self.assertIn("must match", res["errors"][0]["error"])

    def test_bad_code_rejected(self):
        res = self.run_upload([self.row(cmt_codes="01, 100")])
        self.assertEqual(len(res["errors"]), 1)

    def test_duplicate_sequence_inside_file(self):
        res = self.run_upload([self.row(), self.row(attribute_name="RATING", possible_values="150", cmt_codes="01")])
        self.assertEqual(len(res["errors"]), 1)
        self.assertEqual(res["errors"][0]["row"], 3)
        self.assertIn("twice in this file", res["errors"][0]["error"])
        self.assertEqual(MatgAttributeItem.objects.count(), 1)

    def test_duplicate_sequence_against_database(self):
        self.add_char("OLD", 10, ["1"], {"1": "01"})
        res = self.run_upload([self.row()])
        self.assertIn("already assigned", res["errors"][0]["error"])

    def test_missing_sequence_rejected(self):
        res = self.run_upload([self.row(print_priority=None)])
        self.assertIn("required", res["errors"][0]["error"])

    def test_template_has_cmt_column(self):
        from uploads.views import generate_matgattribute_template
        import io, openpyxl
        resp = generate_matgattribute_template(MatgAttributeItem)
        ws = openpyxl.load_workbook(io.BytesIO(resp.content)).active
        headers = [c.value for c in ws[1]]
        self.assertIn("Cmt Codes", headers)
