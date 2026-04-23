# -*- coding: utf-8 -*-
"""Split Sheathing - Calculate sheathing panel layout.

Divides a wall surface into standard 4'x8' sheathing panels
and reports the cut list.
"""

import os
import sys
import math

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

from wf_geometry import analyze_wall, inches_to_feet

output = script.get_output()


# Standard sheathing panel dimensions
PANEL_WIDTH = 4.0   # feet
PANEL_HEIGHT = 8.0  # feet


def main():
    doc = revit.doc

    selected = revit.get_selection().elements
    walls = [e for e in selected if isinstance(e, DB.Wall)]

    if not walls:
        forms.alert(
            "Select walls to calculate sheathing layout.",
            title="Split Sheathing",
        )
        return

    output.print_md("## Sheathing Layout")

    total_full = 0
    total_partial = 0

    for wall in walls:
        wall_info = analyze_wall(doc, wall)
        if wall_info is None:
            continue

        wall_length = wall_info.length
        wall_height = wall_info.height

        # Calculate horizontal panel count
        h_count = int(math.ceil(wall_length / PANEL_WIDTH))
        # Calculate vertical panel count
        v_count = int(math.ceil(wall_height / PANEL_HEIGHT))

        full_panels = 0
        partial_panels = 0

        for row in range(v_count):
            remaining_height = wall_height - row * PANEL_HEIGHT
            panel_h = min(PANEL_HEIGHT, remaining_height)

            for col in range(h_count):
                remaining_width = wall_length - col * PANEL_WIDTH
                panel_w = min(PANEL_WIDTH, remaining_width)

                if abs(panel_w - PANEL_WIDTH) < 0.01 and abs(panel_h - PANEL_HEIGHT) < 0.01:
                    full_panels += 1
                else:
                    partial_panels += 1

        total_full += full_panels
        total_partial += partial_panels

        output.print_md(
            "**Wall {0}:** {1:.1f}' x {2:.1f}' = "
            "{3} full panels + {4} cut panels".format(
                wall.Id.Value,
                wall_length, wall_height,
                full_panels, partial_panels,
            )
        )

    output.print_md(
        "\n---\n**Totals:** {0} full panels, {1} cut panels, "
        "{2} total sheets needed".format(
            total_full, total_partial, total_full + total_partial
        )
    )


if __name__ == "__main__":
    main()
