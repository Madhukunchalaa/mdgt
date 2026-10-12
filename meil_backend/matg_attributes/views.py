from django.http import JsonResponse
from django.views.decorators.csrf import csrf_exempt
from django.utils.decorators import method_decorator
from django.utils import timezone
from django.db import IntegrityError, transaction
import json

from .models import MatgAttributeItem
from matgroups.models import MatGroup
from Employee.models import Employee
from itemmaster.models import ItemMaster
from Common.Middleware import authenticate, restrict
from . import catalog


# Helper function to get employee name
def get_employee_name(emp):
    return emp.emp_name if emp else None


# ============================================================
# ✅ CREATE OR UPDATE MULTIPLE ATTRIBUTE ROWS
# ============================================================
@csrf_exempt
@authenticate
# @restrict(roles=["Admin", "SuperAdmin", "MDGT"])
def create_matgattribute(request):
    if request.method != "POST":
        return JsonResponse({"error": "Invalid request method"}, status=405)

    try:
        data = json.loads(request.body.decode("utf-8"))

        mgrp_code = data.get("mgrp_code")
        attributes = data.get("attributes")  # expecting list

        if not mgrp_code or not isinstance(attributes, list):
            return JsonResponse({"error": "Invalid payload (mgrp_code, attributes)"}, status=400)

        # Validate MatGroup
        matgroup = MatGroup.objects.filter(mgrp_code=mgrp_code, is_deleted=False).first()
        if not matgroup:
            return JsonResponse({"error": "MatGroup not found"}, status=404)

        employee = Employee.objects.filter(emp_id=request.user.get("emp_id")).first()

        # ---------- PASS 1: validate everything, save nothing ----------
        max_seq = MatgAttributeItem.objects.filter(
            mgrp_code=matgroup, is_deleted=False
        ).order_by("-print_priority").values_list("print_priority", flat=True).first() or 0

        prepared = []
        used_in_request = {}    # sequence -> attribute name, for rows in THIS request
        names_in_request = set()

        for attr in attributes:
            attribute_name = (attr.get("attribute_name") or "").strip().upper()
            possible_values = attr.get("possible_values", [])
            if not attribute_name or not isinstance(possible_values, list):
                return JsonResponse({"error": "Invalid attribute structure"}, status=400)

            if attribute_name in names_in_request:
                return JsonResponse({"error": f"Attribute '{attribute_name}' is listed twice in this request"}, status=400)
            names_in_request.add(attribute_name)

            # Already exists (even if deleted)?
            if MatgAttributeItem.objects.filter(mgrp_code=matgroup, attribute_name=attribute_name).exists():
                return JsonResponse({
                    "error": f"Data is already loaded for attribute '{attribute_name}' in this material group. It may be in the 'Deleted' list."
                }, status=400)

            # ---- Sequence: auto-assign if missing, otherwise must be a positive whole number
            raw_seq = attr.get("print_priority") or attr.get("sequence")
            if raw_seq in (None, "", 0, "0"):
                max_seq += 10
                sequence = max_seq
            else:
                try:
                    sequence = catalog.parse_sequence(raw_seq)
                except ValueError as exc:
                    return JsonResponse({"error": f"{attribute_name}: {exc}"}, status=400)

            clash = catalog.find_sequence_conflict(matgroup, sequence)
            if clash:
                return JsonResponse({
                    "error": f"Sequence {sequence} is already assigned to '{clash.attribute_name}' in this material group. Please use a different sequence.",
                    "code": "DUPLICATE_SEQUENCE",
                }, status=400)
            if sequence in used_in_request:
                return JsonResponse({
                    "error": f"Sequence {sequence} is used by both '{used_in_request[sequence]}' and '{attribute_name}'. Please use a different sequence.",
                    "code": "DUPLICATE_SEQUENCE",
                }, status=400)
            used_in_request[sequence] = attribute_name

            # ---- CMT codes: mandatory for every value, exactly 2 digits
            clean_codes, code_errors = catalog.validate_cmt_codes(possible_values, attr.get("cmt_codes"))
            if code_errors:
                return JsonResponse({
                    "error": f"{attribute_name}: " + "; ".join(code_errors),
                    "code": "INVALID_CMT_CODE",
                    "details": code_errors,
                }, status=400)

            prepared.append(dict(
                attribute_name=attribute_name, possible_values=possible_values,
                cmt_codes=clean_codes, uom=attr.get("uom"), sequence=sequence,
                validation=attr.get("validation"),
            ))

        # ---------- PASS 2: save (all or nothing) ----------
        created_items = []
        with transaction.atomic():
            for p in prepared:
                item = MatgAttributeItem.objects.create(
                    mgrp_code=matgroup,
                    attribute_name=p["attribute_name"],
                    possible_values=p["possible_values"],
                    cmt_codes=p["cmt_codes"],
                    uom=p["uom"],
                    print_priority=p["sequence"],
                    validation=p["validation"],
                    createdby=employee,
                    updatedby=employee,
                    updated=timezone.now(),
                    is_deleted=False
                )
                created_items.append({
                    "id": item.id,
                    "attribute_name": item.attribute_name,
                    "possible_values": item.possible_values,
                    "cmt_codes": item.cmt_codes,
                    "uom": item.uom,
                    "sequence": item.print_priority,
                    "print_priority": item.print_priority,
                })

        return JsonResponse({
            "message": "Attributes created/updated successfully",
            "attributes": created_items
        }, status=201)

    except json.JSONDecodeError:
        return JsonResponse({"error": "Invalid JSON"}, status=400)
    except IntegrityError:
        return JsonResponse({"error": "A duplicate attribute record already exists for this material group"}, status=400)



# ============================================================
# ✅ LIST ALL ATTRIBUTE ROWS
# ============================================================
@authenticate
# @restrict(roles=["Admin", "SuperAdmin", "User", "MDGT"])
def list_matgattributes(request):
    if request.method != "GET":
        return JsonResponse({"error": "Invalid request method"}, status=405)

    include_deleted = request.GET.get('include_deleted', 'false').lower() == 'true'
    if include_deleted:
        items = MatgAttributeItem.objects.all()
    else:
        items = MatgAttributeItem.objects.filter(is_deleted=False)

    mgrp_code_filter = request.GET.get('mgrp_code')
    if mgrp_code_filter:
        items = items.filter(mgrp_code__mgrp_code=mgrp_code_filter.strip().upper())

    items = items.order_by('print_priority')

    data = []
    for item in items:
        data.append({
            "id": item.id,
            "mgrp_code": item.mgrp_code.mgrp_code if item.mgrp_code else None,
            "attribute_name": item.attribute_name,
            "possible_values": item.possible_values,
            "cmt_codes": item.cmt_codes or {},
            "uom": item.uom,
            "validation": item.validation,
            "sequence": item.print_priority,
            "print_priority": item.print_priority,
            "is_deleted": item.is_deleted,
            "created": item.created.strftime("%Y-%m-%d %H:%M:%S"),
            "updated": item.updated.strftime("%Y-%m-%d %H:%M:%S"),
            "createdby": get_employee_name(item.createdby),
            "updatedby": get_employee_name(item.updatedby),
        })

    return JsonResponse(data, safe=False)



# ============================================================
# ✅ UPDATE A SINGLE ATTRIBUTE ROW
# ============================================================
@csrf_exempt
@authenticate
# @restrict(roles=["Admin", "SuperAdmin", "MDGT"])
def update_matgattribute(request, item_id):
    if request.method != "PUT":
        return JsonResponse({"error": "Invalid request method"}, status=405)

    try:
        data = json.loads(request.body.decode("utf-8"))

        item = MatgAttributeItem.objects.filter(id=item_id, is_deleted=False).first()
        if not item:
            return JsonResponse({"error": "Attribute item not found"}, status=404)

        employee = Employee.objects.filter(emp_id=request.user.get("emp_id")).first()
        affects_catalog = False

        # Update fields
        if "attribute_name" in data:
            item.attribute_name = data["attribute_name"].strip().upper()

        values_or_codes_changed = "possible_values" in data or "cmt_codes" in data
        if "possible_values" in data:
            if not isinstance(data["possible_values"], list):
                return JsonResponse({"error": "possible_values must be a list"}, status=400)
            item.possible_values = data["possible_values"]

        if values_or_codes_changed:
            # Codes sent now win; otherwise keep the ones already stored
            codes_in = data["cmt_codes"] if "cmt_codes" in data else item.cmt_codes
            clean_codes, code_errors = catalog.validate_cmt_codes(item.possible_values, codes_in)
            if code_errors:
                return JsonResponse({
                    "error": f"{item.attribute_name}: " + "; ".join(code_errors),
                    "code": "INVALID_CMT_CODE",
                    "details": code_errors,
                }, status=400)
            item.cmt_codes = clean_codes
            affects_catalog = True

        if "uom" in data:
            item.uom = data["uom"]
            affects_catalog = True

        if "print_priority" in data or "sequence" in data:
            raw_seq = data.get("sequence") if data.get("sequence") not in (None, "") else data.get("print_priority")
            blank = raw_seq in (None, "", 0, "0")
            if blank and item.print_priority is None:
                pass    # old characteristic that never had a sequence: leave it alone, don't block other edits
            else:
                try:
                    new_priority = catalog.parse_sequence(raw_seq)
                except ValueError as exc:
                    return JsonResponse({"error": str(exc)}, status=400)
                clash = catalog.find_sequence_conflict(item.mgrp_code, new_priority, exclude_id=item.id)
                if clash:
                    return JsonResponse({
                        "error": f"Sequence {new_priority} is already assigned to '{clash.attribute_name}' in this material group. Please use a different sequence.",
                        "code": "DUPLICATE_SEQUENCE",
                    }, status=400)
                if new_priority != item.print_priority:
                    affects_catalog = True
                item.print_priority = new_priority

        if "validation" in data:
            item.validation = data["validation"] if data["validation"] else None

        item.updatedby = employee
        item.updated = timezone.now()
        with transaction.atomic():
            item.save()
            refreshed = catalog.refresh_group_catalog_numbers(item.mgrp_code) if affects_catalog else 0

        return JsonResponse({
            "message": "Attribute updated successfully",
            "catalog_numbers_refreshed": refreshed,
        }, status=200)

    except json.JSONDecodeError:
        return JsonResponse({"error": "Invalid JSON"}, status=400)



# ============================================================
# ✅ DELETE A SINGLE ATTRIBUTE ROW
# ============================================================
@csrf_exempt
@authenticate
# @restrict(roles=["Admin", "SuperAdmin", "MDGT"])
def delete_matgattribute(request, item_id):
    if request.method != "DELETE":
        return JsonResponse({"error": "Invalid request method"}, status=405)

    item = MatgAttributeItem.objects.filter(id=item_id).first()
    if not item:
        return JsonResponse({"error": "Attribute item not found"}, status=404)

    # Block deletion if this attribute is used by items in the same material group
    used_count = ItemMaster.objects.filter(
        is_deleted=False,
        mgrp_code=item.mgrp_code,
        attributes__has_key=item.attribute_name
    ).count()
    if used_count > 0:
        return JsonResponse({
            "error": f"Cannot delete '{item.attribute_name}'. It is used by {used_count} item(s) in this material group. Remove it from all items first."
        }, status=400)

    item.is_deleted = True
    item.save()
    catalog.refresh_group_catalog_numbers(item.mgrp_code)

    return JsonResponse({"message": "Attribute deleted successfully"}, status=200)


@csrf_exempt
@authenticate
def restore_matgattribute(request, item_id):
    if request.method != "POST":
        return JsonResponse({"error": "Invalid request method"}, status=405)

    item = MatgAttributeItem.objects.filter(id=item_id).first()
    if not item:
        return JsonResponse({"error": "Attribute item not found"}, status=404)

    if item.print_priority:
        clash = catalog.find_sequence_conflict(item.mgrp_code, item.print_priority, exclude_id=item.id)
        if clash:
            return JsonResponse({
                "error": f"Cannot restore: sequence {item.print_priority} is now used by '{clash.attribute_name}'. Change the sequence of one of them first.",
                "code": "DUPLICATE_SEQUENCE",
            }, status=400)

    item.is_deleted = False
    item.save()
    catalog.refresh_group_catalog_numbers(item.mgrp_code)

    return JsonResponse({"message": "Attribute restored successfully"}, status=200)



# ============================================================
# 🆕 PREVIEW: what Catalog Number would this selection produce?
# ============================================================
@csrf_exempt
@authenticate
def preview_catalog_number(request):
    """POST {mgrp_code, attributes:{NAME: value}} -> catalog number + per-characteristic breakdown."""
    if request.method != "POST":
        return JsonResponse({"error": "Invalid request method"}, status=405)
    try:
        data = json.loads(request.body.decode("utf-8"))
    except json.JSONDecodeError:
        return JsonResponse({"error": "Invalid JSON"}, status=400)

    matgroup = MatGroup.objects.filter(mgrp_code=data.get("mgrp_code"), is_deleted=False).first()
    if not matgroup:
        return JsonResponse({"error": "MatGroup not found"}, status=404)
    return JsonResponse(catalog.catalog_breakdown(matgroup, data.get("attributes") or {}))


# ============================================================
# 🆕 DATA CLEAN-UP: which material groups still have duplicate sequences?
# ============================================================
@authenticate
def sequence_conflicts(request):
    if request.method != "GET":
        return JsonResponse({"error": "Invalid request method"}, status=405)
    mgrp = request.GET.get("mgrp_code")
    group = MatGroup.objects.filter(mgrp_code=mgrp.strip().upper()).first() if mgrp else None
    conflicts = catalog.duplicate_sequences(group)
    return JsonResponse({"count": len(conflicts), "conflicts": conflicts})
