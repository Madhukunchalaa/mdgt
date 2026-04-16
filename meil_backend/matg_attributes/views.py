from django.http import JsonResponse
from django.views.decorators.csrf import csrf_exempt
from django.utils.decorators import method_decorator
from django.utils import timezone
from django.db import IntegrityError
import json

from .models import MatgAttributeItem
from matgroups.models import MatGroup
from Employee.models import Employee
from itemmaster.models import ItemMaster
from Common.Middleware import authenticate, restrict


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

        created_items = []

        # Loop and create/update each attribute row
        for attr in attributes:
            attribute_name = attr.get("attribute_name", "").strip().upper()
            possible_values = attr.get("possible_values", [])
            uom = attr.get("uom")
            print_priority = attr.get("print_priority") or attr.get("sequence")
            validation = attr.get("validation")

            if not attribute_name or not isinstance(possible_values, list):
                return JsonResponse({"error": "Invalid attribute structure"}, status=400)

            # Check if this attribute already exists (even if deleted)
            existing = MatgAttributeItem.objects.filter(
                mgrp_code=matgroup,
                attribute_name=attribute_name
            ).first()
            if existing:
                return JsonResponse({
                    "error": f"Data is already loaded for attribute '{attribute_name}' in this material group. It may be in the 'Deleted' list."
                }, status=400)

            # Auto-assign sequence if not provided: max existing + 10, starting at 10
            if not print_priority:
                max_seq = MatgAttributeItem.objects.filter(
                    mgrp_code=matgroup, is_deleted=False
                ).exclude(attribute_name=attribute_name).order_by("-print_priority").values_list("print_priority", flat=True).first()
                print_priority = (max_seq or 0) + 10

            # Check for duplicate sequence (non-zero)
            if print_priority:
                duplicate = MatgAttributeItem.objects.filter(
                    mgrp_code=matgroup,
                    print_priority=print_priority,
                    is_deleted=False
                ).exclude(attribute_name=attribute_name).first()
                if duplicate:
                    return JsonResponse({
                        "error": f"Sequence {print_priority} is already assigned to attribute '{duplicate.attribute_name}' in this material group"
                    }, status=400)

            item = MatgAttributeItem.objects.create(
                mgrp_code=matgroup,
                attribute_name=attribute_name,
                possible_values=possible_values,
                uom=uom,
                print_priority=print_priority,
                validation=validation,
                createdby=employee,
                updatedby=employee,
                updated=timezone.now(),
                is_deleted=False
            )

            created_items.append({
                "id": item.id,
                "attribute_name": item.attribute_name,
                "possible_values": item.possible_values,
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

        # Update fields
        if "attribute_name" in data:
            item.attribute_name = data["attribute_name"].strip().upper()

        if "possible_values" in data:
            if not isinstance(data["possible_values"], list):
                return JsonResponse({"error": "possible_values must be a list"}, status=400)
            item.possible_values = data["possible_values"]

        if "uom" in data:
            item.uom = data["uom"]

        if "print_priority" in data or "sequence" in data:
            new_priority = data.get("sequence") or data.get("print_priority")
            if new_priority:
                duplicate = MatgAttributeItem.objects.filter(
                    mgrp_code=item.mgrp_code,
                    print_priority=new_priority,
                    is_deleted=False
                ).exclude(id=item_id).first()
                if duplicate:
                    return JsonResponse({
                        "error": f"Sequence {new_priority} is already assigned to attribute '{duplicate.attribute_name}' in this material group"
                    }, status=400)
            item.print_priority = new_priority

        if "validation" in data:
            item.validation = data["validation"] if data["validation"] else None

        item.updatedby = employee
        item.updated = timezone.now()
        item.save()

        return JsonResponse({"message": "Attribute updated successfully"}, status=200)

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


    return JsonResponse({"message": "Attribute deleted successfully"}, status=200)


@csrf_exempt
@authenticate
def restore_matgattribute(request, item_id):
    if request.method != "POST":
        return JsonResponse({"error": "Invalid request method"}, status=405)

    item = MatgAttributeItem.objects.filter(id=item_id).first()
    if not item:
        return JsonResponse({"error": "Attribute item not found"}, status=404)

    item.is_deleted = False
    item.save()

    return JsonResponse({"message": "Attribute restored successfully"}, status=200)
