# -*- coding: utf-8 -*-
"""Frame Floor - Main command script.

Select floors and automatically generate floor framing:
joists, rim joists, and blocking.
"""

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

from wf_config import FramingConfig, SPACING_16OC, SPACING_24OC
from wf_families import get_available_types_flat, parse_family_type_label
from wf_floor import FloorFramingEngine

logger = script.get_logger()
output = script.get_output()


def main():
    doc = revit.doc

    available_types = get_available_types_flat(doc)
    if not available_types:
        forms.alert(
            "No structural framing families are loaded.\n"
            "Load a framing family before running this command.",
            title="Wood Framing",
        )
        return

    # Select floors
    selected = revit.get_selection().elements
    floors = [e for e in selected if isinstance(e, DB.Floor)]

    if not floors:
        try:
            refs = revit.uidoc.Selection.PickObjects(
                ObjectType.Element,
                _FloorFilter(),
                "Select floors to frame",
            )
            floors = [doc.GetElement(r.ElementId) for r in refs]
        except Exception:
            return

    if not floors:
        forms.alert("No floors selected.", title="Wood Framing")
        return

    # Quick config dialog
    joist_type = forms.SelectFromList.show(
        available_types,
        title="Select Joist Type",
        button_name="OK",
    )
    if not joist_type:
        return

    spacing_choice = forms.CommandSwitchWindow.show(
        ['16" OC', '24" OC'],
        message="Select joist spacing:",
    )
    if not spacing_choice:
        return

    cfg = FramingConfig()
    cfg.stud_spacing = SPACING_16OC if '16' in spacing_choice else SPACING_24OC
    fam, typ = parse_family_type_label(str(joist_type))
    cfg.stud_family_name = fam
    cfg.stud_type_name = typ
    cfg.bottom_plate_family_name = fam
    cfg.bottom_plate_type_name = typ

    engine = FloorFramingEngine(doc, cfg)
    total_placed = 0
    total_floors = 0

    with revit.Transaction("WF: Frame Floors"):
        for floor in floors:
            members, floor_info = engine.calculate_members(floor)
            if floor_info is None:
                logger.warning(
                    "Skipped floor {0}.".format(floor.Id.Value)
                )
                continue
            placed = engine.place_members(members, floor_info)
            total_placed += len(placed)
            total_floors += 1

    output.print_md(
        "## Floor Framing Complete\n"
        "- **Floors framed:** {0}\n"
        "- **Members placed:** {1}".format(total_floors, total_placed)
    )


class _FloorFilter(ISelectionFilter):
    def AllowElement(self, element):
        return isinstance(element, DB.Floor)

    def AllowReference(self, reference, point):
        return False


if __name__ == "__main__":
    main()
