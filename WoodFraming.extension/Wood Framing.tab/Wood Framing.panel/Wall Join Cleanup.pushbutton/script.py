# -*- coding: utf-8 -*-
"""Wall Join Cleanup command."""

import json
import os
import sys

from pyrevit import DB, forms, revit, script
from pyrevit.forms import WPFWindow
from Autodesk.Revit.UI.Selection import ISelectionFilter, ObjectType


_ext_dir = os.path.dirname(os.path.dirname(os.path.dirname(os.path.dirname(__file__))))
while _ext_dir and not _ext_dir.lower().endswith(".extension"):
    _parent = os.path.dirname(_ext_dir)
    if _parent == _ext_dir:
        break
    _ext_dir = _parent
_lib_dir = os.path.join(_ext_dir, "lib")
if _lib_dir not in sys.path:
    sys.path.insert(0, _lib_dir)

from wf_config import FramingConfig
from wf_families import (
    get_available_types_flat,
    get_column_types_flat,
    parse_family_type_label,
)
from wf_wall_join_cleanup import (
    JOIN_KIND_CORNER,
    JOIN_KIND_T,
    STYLE_CORNER_CAVITY,
    STYLE_CORNER_INSULATED,
    STYLE_T_ASSEMBLY,
    STYLE_T_BLOCKING_NAILER,
    WallJoinCleanupError,
    analyze_wall_join,
    cleanup_selected_wall_join,
)


logger = script.get_logger()
output = script.get_output()
COMMAND_TITLE = "Wall Join Cleanup"
_XAML = os.path.join(os.path.dirname(__file__), "FrameWallJoinCleanupConfig.xaml")
_CFG_PATH = os.path.join(
    os.environ.get("APPDATA", ""),
    "pyRevit",
    "WoodFraming_WallJoinCleanupLastConfig.json",
)

CORNER_STYLE_LABELS = [
    ("Insulated corner - two studs", STYLE_CORNER_INSULATED),
    ("Cavity corner - two studs plus blocking", STYLE_CORNER_CAVITY),
]
T_STYLE_LABELS = [
    ("T assembly", STYLE_T_ASSEMBLY),
    ("Blocking and 1x nailer", STYLE_T_BLOCKING_NAILER),
]


class WallJoinCleanupDialog(WPFWindow):
    def __init__(self, column_types, framing_types, relation):
        WPFWindow.__init__(self, _XAML)
        self.result_config = None
        self.result_style_key = None
        self._column_types = column_types
        self._framing_types = framing_types
        self._relation = relation
        self._style_pairs = _style_pairs_for_join(relation.kind)

        self.cb_stud_type.ItemsSource = column_types
        self.cb_blocking_type.ItemsSource = framing_types
        self.cb_assembly.ItemsSource = [label for label, _key in self._style_pairs]

        if column_types:
            self.cb_stud_type.SelectedIndex = 0
        if framing_types:
            self.cb_blocking_type.SelectedIndex = 0
        if self._style_pairs:
            self.cb_assembly.SelectedIndex = 0

        self.tb_summary.Text = (
            "{0} detected between the two selected walls. Angle: {1:.1f} degrees."
            .format(_join_label(relation.kind), relation.angle_degrees)
        )
        self._restore_last()

    def btn_ok_click(self, sender, args):
        config = self._build_config()
        if config is None:
            return
        style_key = self._selected_style_key()
        if style_key is None:
            forms.alert("Select a join assembly.", title=COMMAND_TITLE)
            return
        self.result_config = config
        self.result_style_key = style_key
        self._save_last()
        self.Close()

    def btn_cancel_click(self, sender, args):
        self.result_config = None
        self.result_style_key = None
        self.Close()

    def _build_config(self):
        stud_label = self.cb_stud_type.SelectedItem
        if not stud_label:
            forms.alert("Select a stud column family.", title=COMMAND_TITLE)
            return None

        blocking_label = self.cb_blocking_type.SelectedItem
        if not blocking_label:
            forms.alert("Select a blocking/nailer framing family.", title=COMMAND_TITLE)
            return None

        config = FramingConfig()
        config.track_members = True
        config.stud_family_name, config.stud_type_name = parse_family_type_label(
            str(stud_label)
        )
        family, type_name = parse_family_type_label(str(blocking_label))
        config.bottom_plate_family_name = family
        config.bottom_plate_type_name = type_name
        config.top_plate_family_name = family
        config.top_plate_type_name = type_name
        return config

    def _selected_style_key(self):
        selected = self.cb_assembly.SelectedItem
        if not selected:
            return None
        for label, key in self._style_pairs:
            if str(selected) == label:
                return key
        return None

    def _save_last(self):
        data = {}
        try:
            if os.path.exists(_CFG_PATH):
                with open(_CFG_PATH, "r") as cfg_file:
                    data = json.load(cfg_file)
        except Exception:
            data = {}

        data.update({
            "stud": str(self.cb_stud_type.SelectedItem or ""),
            "blocking": str(self.cb_blocking_type.SelectedItem or ""),
        })
        selected_assembly = str(self.cb_assembly.SelectedItem or "")
        if self._relation.kind == JOIN_KIND_CORNER:
            data["corner_assembly"] = selected_assembly
        elif self._relation.kind == JOIN_KIND_T:
            data["t_assembly"] = selected_assembly

        try:
            cfg_dir = os.path.dirname(_CFG_PATH)
            if not os.path.exists(cfg_dir):
                os.makedirs(cfg_dir)
            with open(_CFG_PATH, "w") as cfg_file:
                json.dump(data, cfg_file, indent=2)
        except Exception:
            pass

    def _restore_last(self):
        try:
            if not os.path.exists(_CFG_PATH):
                return
            with open(_CFG_PATH, "r") as cfg_file:
                data = json.load(cfg_file)
        except Exception:
            return

        self._select_if_present(self.cb_stud_type, data.get("stud"))
        self._select_if_present(self.cb_blocking_type, data.get("blocking"))
        if self._relation.kind == JOIN_KIND_CORNER:
            self._select_if_present(self.cb_assembly, data.get("corner_assembly"))
        elif self._relation.kind == JOIN_KIND_T:
            self._select_if_present(self.cb_assembly, data.get("t_assembly"))

    @staticmethod
    def _select_if_present(combo, value):
        if not value:
            return
        try:
            for item in combo.Items:
                if str(item) == str(value):
                    combo.SelectedItem = item
                    return
        except Exception:
            pass


class _WallFilter(ISelectionFilter):
    def AllowElement(self, element):
        return isinstance(element, DB.Wall)

    def AllowReference(self, reference, point):
        return False


def _selected_or_pick_two_walls(doc):
    selected = revit.get_selection().elements
    walls = [element for element in selected if isinstance(element, DB.Wall)]
    if len(walls) == 2:
        return walls

    try:
        refs = revit.uidoc.Selection.PickObjects(
            ObjectType.Element,
            _WallFilter(),
            "Select exactly two walls for corner/T cleanup",
        )
    except Exception:
        return []

    walls = [doc.GetElement(ref.ElementId) for ref in refs]
    walls = [wall for wall in walls if isinstance(wall, DB.Wall)]
    if len(walls) != 2:
        forms.alert("Select exactly two walls.", title=COMMAND_TITLE)
        return []
    return walls


def _style_pairs_for_join(join_kind):
    if join_kind == JOIN_KIND_CORNER:
        return CORNER_STYLE_LABELS
    if join_kind == JOIN_KIND_T:
        return T_STYLE_LABELS
    return []


def _style_label(style_key):
    for label, key in CORNER_STYLE_LABELS + T_STYLE_LABELS:
        if key == style_key:
            return label
    return style_key or ""


def _join_label(join_kind):
    if join_kind == JOIN_KIND_CORNER:
        return "Corner"
    if join_kind == JOIN_KIND_T:
        return "T intersection"
    return join_kind or ""


def main():
    doc = revit.doc

    column_types = get_column_types_flat(doc)
    framing_types = get_available_types_flat(doc)
    if not column_types:
        forms.alert("No Structural Column families loaded.", title=COMMAND_TITLE)
        return
    if not framing_types:
        forms.alert("No Structural Framing families loaded.", title=COMMAND_TITLE)
        return

    walls = _selected_or_pick_two_walls(doc)
    if not walls:
        return

    try:
        relation = analyze_wall_join(doc, walls, FramingConfig())
    except WallJoinCleanupError as exc:
        forms.alert(str(exc), title=COMMAND_TITLE)
        return

    dialog = WallJoinCleanupDialog(column_types, framing_types, relation)
    dialog.ShowDialog()
    config = dialog.result_config
    style_key = dialog.result_style_key
    if config is None or style_key is None:
        return

    try:
        with revit.Transaction("WF: Wall Join Cleanup"):
            result = cleanup_selected_wall_join(doc, walls, config, style_key)
    except WallJoinCleanupError as exc:
        forms.alert(str(exc), title=COMMAND_TITLE)
        return

    warning_text = ""
    if result.warnings:
        warning_text = "\n".join(["- **Warning:** {0}".format(item) for item in result.warnings])
        warning_text = "\n{0}\n".format(warning_text)

    output.print_md(
        "## Wall Join Cleanup Complete\n"
        "- **Join type:** {0}\n"
        "- **Assembly:** {1}\n"
        "- **Angle:** {2:.1f} degrees\n"
        "- **Previous join assembly members replaced:** {3}\n"
        "- **Assembly members requested:** {4}\n"
        "- **Assembly members placed:** {5}\n"
        "{6}".format(
            _join_label(result.join_kind),
            _style_label(result.style_key),
            result.angle_degrees,
            result.deleted_count,
            result.requested_count,
            result.placed_count,
            warning_text,
        )
    )


if __name__ == "__main__":
    main()
