# -*- coding: utf-8 -*-
"""Generate BOM - Create a Bill of Materials for generated wood framing.

Collects generated structural framing and structural column members and
outputs a summary grouped by member role and family/type.
"""

import os
import sys
from collections import OrderedDict

_ext_dir = os.path.dirname(os.path.dirname(os.path.dirname(
    os.path.dirname(__file__)
)))
while _ext_dir and not _ext_dir.lower().endswith(".extension"):
    _parent = os.path.dirname(_ext_dir)
    if _parent == _ext_dir:
        break
    _ext_dir = _parent
_lib_dir = os.path.join(_ext_dir, "lib")
if _lib_dir not in sys.path:
    sys.path.insert(0, _lib_dir)

from pyrevit import revit, DB, script, forms
from wf_tracking import get_tracking_data

output = script.get_output()


def main():
    doc = revit.doc

    from Autodesk.Revit.DB import (
        BuiltInCategory, BuiltInParameter, Element as _Element,
    )

    elements = _collect_generated_elements(doc)
    if not elements:
        forms.alert("No generated framing members found in the project.",
                     title="Generate BOM")
        return

    bom = OrderedDict()
    for elem, member_role in elements:
        try:
            fam_name = _Element.Name.__get__(elem.Symbol.Family)
            type_name = elem.Symbol.get_Parameter(
                BuiltInParameter.ALL_MODEL_TYPE_NAME
            ).AsString() or _Element.Name.__get__(elem)
        except Exception:
            fam_name = "Unknown"
            type_name = "Unknown"

        key = (member_role, "{0} : {1}".format(fam_name, type_name))
        length_ft = _element_length_ft(elem)

        if key not in bom:
            bom[key] = {
                "member": member_role,
                "family_type": key[1],
                "count": 0,
                "total_length_ft": 0.0,
            }
        bom[key]["count"] += 1
        bom[key]["total_length_ft"] += length_ft

    # Output as formatted table
    output.print_md("## Wood Framing - Bill of Materials")
    output.print_md("---")

    table_data = []
    total_count = 0
    total_length = 0.0

    for key in sorted(bom.keys()):
        data = bom[key]
        count = data["count"]
        length_ft = data["total_length_ft"]
        # Convert feet to feet-inches display
        feet = int(length_ft)
        inches = (length_ft - feet) * 12.0
        length_str = "{0}'-{1:.1f}\"".format(feet, inches)

        table_data.append([
            data["member"],
            data["family_type"],
            str(count),
            length_str,
        ])
        total_count += count
        total_length += length_ft

    output.print_table(
        table_data,
        title="Framing Members",
        columns=["Member", "Family : Type", "Count", "Total Length"],
    )

    total_ft = int(total_length)
    total_in = (total_length - total_ft) * 12.0
    output.print_md(
        "\n**Totals:** {0} members, {1}'-{2:.1f}\" total length".format(
            total_count, total_ft, total_in
        )
    )


def _collect_generated_elements(doc):
    from Autodesk.Revit.DB import FilteredElementCollector, BuiltInCategory

    elements = []
    seen_ids = set()
    categories = (
        BuiltInCategory.OST_StructuralFraming,
        BuiltInCategory.OST_StructuralColumns,
    )

    for category in categories:
        collector = (
            FilteredElementCollector(doc)
            .OfCategory(category)
            .WhereElementIsNotElementType()
        )
        for elem in collector:
            elem_id = getattr(elem.Id, "IntegerValue", getattr(elem.Id, "Value", None))
            if elem_id in seen_ids:
                continue

            member_role = _member_role(elem)
            if member_role is None:
                continue

            seen_ids.add(elem_id)
            elements.append((elem, member_role))

    return elements


def _member_role(elem):
    tracking = get_tracking_data(elem)
    if tracking is not None:
        return tracking.get("member") or "generated"

    comments = elem.get_Parameter(DB.BuiltInParameter.ALL_MODEL_INSTANCE_COMMENTS)
    if comments is None:
        return None

    text = comments.AsString() or ""
    if "WF_Generated" in text:
        return "generated"
    return None


def _element_length_ft(elem):
    for bip in (
        DB.BuiltInParameter.STRUCTURAL_FRAME_CUT_LENGTH,
        DB.BuiltInParameter.INSTANCE_LENGTH_PARAM,
    ):
        param = elem.get_Parameter(bip)
        if param is not None and param.HasValue:
            value = param.AsDouble()
            if value > 0.0:
                return value

    loc = getattr(elem, "Location", None)
    curve = getattr(loc, "Curve", None)
    if curve is not None:
        try:
            return curve.Length
        except Exception:
            pass

    bbox = elem.get_BoundingBox(None)
    if bbox is not None:
        return max(
            abs(bbox.Max.X - bbox.Min.X),
            abs(bbox.Max.Y - bbox.Min.Y),
            abs(bbox.Max.Z - bbox.Min.Z),
        )

    return 0.0


if __name__ == "__main__":
    main()
