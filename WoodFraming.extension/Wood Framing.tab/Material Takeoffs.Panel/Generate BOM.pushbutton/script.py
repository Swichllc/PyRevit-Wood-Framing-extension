# -*- coding: utf-8 -*-
"""Generate the native Revit BOM schedule for wood framing."""

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

from pyrevit import revit, script

from wf_schedule_utils import (
    BOM_SCHEDULE_NAME,
    activate_schedule,
    create_or_update_bom_schedule,
)


output = script.get_output()


def main():
    doc = revit.doc

    with revit.Transaction("WF: Update BOM Schedule"):
        schedule = create_or_update_bom_schedule(doc)

    activate_schedule(schedule)
    output.print_md(
        "## Generate BOM\n"
        "- **Schedule updated:** {0}\n"
        "- **Scope:** generated framing members grouped by host, member role, and family/type".format(
            BOM_SCHEDULE_NAME
        )
    )


if __name__ == "__main__":
    main()
