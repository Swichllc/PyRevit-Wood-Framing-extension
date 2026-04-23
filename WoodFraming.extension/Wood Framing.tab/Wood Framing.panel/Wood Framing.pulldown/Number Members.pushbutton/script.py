# -*- coding: utf-8 -*-
"""Number Members - Assign sequential mark numbers to framing members.

Numbers all structural framing members with a prefix and sequential number,
grouped by family type (e.g. S-001, S-002 for studs, H-001 for headers).
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

output = script.get_output()


def main():
    doc = revit.doc

    from Autodesk.Revit.DB import (
        FilteredElementCollector, BuiltInCategory, BuiltInParameter
    )

    collector = (
        FilteredElementCollector(doc)
        .OfCategory(BuiltInCategory.OST_StructuralFraming)
        .WhereElementIsNotElementType()
    )

    elements = list(collector)
    if not elements:
        forms.alert("No structural framing members found.",
                     title="Number Members")
        return

    # Ask for prefix
    prefix = forms.ask_for_string(
        prompt="Enter mark prefix (e.g. WF):",
        title="Number Members",
        default="WF",
    )
    if prefix is None:
        return

    # Sort by type name, then by location
    from Autodesk.Revit.DB import Element as _Element
    def sort_key(elem):
        try:
            type_name = _Element.Name.__get__(elem.Symbol.Family) + _Element.Name.__get__(elem)
        except Exception:
            type_name = ""
        try:
            loc = elem.Location.Curve.GetEndPoint(0)
            return (type_name, loc.X, loc.Y, loc.Z)
        except Exception:
            return (type_name, 0, 0, 0)

    elements.sort(key=sort_key)

    numbered = 0
    with revit.Transaction("WF: Number Members"):
        for i, elem in enumerate(elements):
            mark = "{0}-{1:04d}".format(prefix, i + 1)
            param = elem.get_Parameter(BuiltInParameter.ALL_MODEL_MARK)
            if param is not None and not param.IsReadOnly:
                param.Set(mark)
                numbered += 1

    output.print_md(
        "## Number Members\n"
        "Assigned marks to **{0}** members.\n"
        "Format: `{1}-NNNN`".format(numbered, prefix)
    )


if __name__ == "__main__":
    main()
