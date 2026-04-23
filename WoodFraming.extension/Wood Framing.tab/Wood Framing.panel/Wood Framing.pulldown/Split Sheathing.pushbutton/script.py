# -*- coding: utf-8 -*-
"""Generate the native Revit sheathing schedule for selected hosts."""

import os
import sys

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
from Autodesk.Revit.UI.Selection import ObjectType, ISelectionFilter

from wf_schedule_utils import (
    SHEATHING_SCHEDULE_NAME,
    activate_schedule,
    calculate_sheathing_for_host,
    clear_all_sheathing_metadata,
    create_or_update_sheathing_schedule,
    ensure_sheathing_parameters,
    stamp_sheathing_metadata,
)


output = script.get_output()


class _SheathingHostFilter(ISelectionFilter):
    def AllowElement(self, element):
        return _is_supported_host(element)

    def AllowReference(self, reference, point):
        return False


def _is_supported_host(element):
    if isinstance(element, DB.Wall):
        return True
    if isinstance(element, DB.Floor):
        return True
    if isinstance(element, DB.RoofBase):
        return True

    category = getattr(element, "Category", None)
    if category is None:
        return False
    category_id = getattr(category.Id, "IntegerValue", getattr(category.Id, "Value", None))
    return category_id == int(DB.BuiltInCategory.OST_Ceilings)


def _select_hosts(doc):
    seen = set()
    hosts = []
    for element in revit.get_selection().elements:
        if not _is_supported_host(element):
            continue
        element_id = getattr(element.Id, "IntegerValue", getattr(element.Id, "Value", None))
        if element_id in seen:
            continue
        seen.add(element_id)
        hosts.append(element)
    if hosts:
        return hosts

    try:
        refs = revit.uidoc.Selection.PickObjects(
            ObjectType.Element,
            _SheathingHostFilter(),
            "Select walls, floors, ceilings, and roofs for sheathing",
        )
    except Exception:
        return []

    for reference in refs:
        element = doc.GetElement(reference.ElementId)
        if not _is_supported_host(element):
            continue
        element_id = getattr(element.Id, "IntegerValue", getattr(element.Id, "Value", None))
        if element_id in seen:
            continue
        seen.add(element_id)
        hosts.append(element)
    return hosts


def main():
    doc = revit.doc
    hosts = _select_hosts(doc)
    if not hosts:
        forms.alert(
            "Select walls, floors, ceilings, or roofs to calculate sheathing.",
            title="Split Sheathing",
        )
        return

    results = []
    with revit.Transaction("WF: Update Sheathing Schedule"):
        ensure_sheathing_parameters(doc)
        clear_all_sheathing_metadata(doc)

        for element in hosts:
            result = calculate_sheathing_for_host(doc, element)
            if result is None:
                continue
            stamp_sheathing_metadata(element, result)
            results.append(result)

        schedule = create_or_update_sheathing_schedule(doc)

    activate_schedule(schedule)

    total_full = sum([result.get("full_sheets", 0) for result in results])
    total_cut = sum([result.get("cut_count", 0) for result in results])
    total_eq = sum([result.get("total_sheet_eq", 0.0) for result in results])

    output.print_md(
        "## Split Sheathing\n"
        "- **Schedule updated:** {0}\n"
        "- **Hosts processed:** {1}\n"
        "- **Full sheets:** {2}\n"
        "- **Cut panels:** {3}\n"
        "- **Total sheet equivalent:** {4:.2f}".format(
            SHEATHING_SCHEDULE_NAME,
            len(results),
            total_full,
            total_cut,
            total_eq,
        )
    )


if __name__ == "__main__":
    main()
