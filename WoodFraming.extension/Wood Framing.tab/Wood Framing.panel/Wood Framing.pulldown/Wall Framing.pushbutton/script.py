# -*- coding: utf-8 -*-
"""Wall Framing command."""

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

from wf_config import (
    FramingConfig,
    SPACING_16OC,
    SPACING_24OC,
    WALL_BASE_MODE_SUPPORT_TOP,
    WALL_BASE_MODE_WALL,
)
from wf_families import (
    get_available_types_flat,
    get_column_types_flat,
    parse_family_type_label,
)
from wf_tracking import get_tracking_data
from wf_wall_framing_v4 import ENGINE_NAME, WallCavityFramingV4Engine


logger = script.get_logger()
output = script.get_output()
COMMAND_TITLE = "Wall Framing"
_XAML = os.path.join(os.path.dirname(__file__), "FrameWall20Config.xaml")
_CFG_PATH = os.path.join(
    os.environ.get("APPDATA", ""),
    "pyRevit",
    "WoodFraming_Wall20LastConfig.json",
)


class WallFraming20Dialog(WPFWindow):
    def __init__(self, column_types, framing_types, wall_count):
        WPFWindow.__init__(self, _XAML)
        self.result_config = None
        self._column_types = column_types
        self._framing_types = framing_types

        self.cb_stud_type.ItemsSource = column_types
        self.cb_plate_type.ItemsSource = framing_types
        self.cb_header_type.ItemsSource = framing_types

        if column_types:
            self.cb_stud_type.SelectedIndex = 0
        if framing_types:
            self.cb_plate_type.SelectedIndex = 0
            self.cb_header_type.SelectedIndex = 0

        self.tb_summary.Text = (
            "Frame {0} selected wall(s)."
            .format(wall_count)
        )

        self.rb_custom.Checked += self._on_custom_checked
        self.rb_custom.Unchecked += self._on_custom_unchecked
        self.rb_16oc.Checked += self._on_custom_unchecked
        self.rb_24oc.Checked += self._on_custom_unchecked
        self.chk_mid_plates.Checked += self._on_mid_plates_checked
        self.chk_mid_plates.Unchecked += self._on_mid_plates_unchecked

        self._restore_last()

    def _on_custom_checked(self, sender, args):
        self.tb_custom_spacing.IsEnabled = True

    def _on_custom_unchecked(self, sender, args):
        self.tb_custom_spacing.IsEnabled = False

    def _on_mid_plates_checked(self, sender, args):
        self.tb_mid_plate_height.IsEnabled = True

    def _on_mid_plates_unchecked(self, sender, args):
        self.tb_mid_plate_height.IsEnabled = False

    def btn_ok_click(self, sender, args):
        config = self._build_config()
        if config is None:
            return
        self.result_config = config
        self._save_last()
        self.Close()

    def btn_cancel_click(self, sender, args):
        self.result_config = None
        self.Close()

    def _build_config(self):
        config = FramingConfig()

        if self.rb_24oc.IsChecked:
            config.stud_spacing = SPACING_24OC
        elif self.rb_custom.IsChecked:
            try:
                config.stud_spacing = float(self.tb_custom_spacing.Text)
            except Exception:
                forms.alert("Invalid custom stud spacing.", title=COMMAND_TITLE)
                return None
        else:
            config.stud_spacing = SPACING_16OC

        stud_sel = self.cb_stud_type.SelectedItem
        if not stud_sel:
            forms.alert("Select a stud column family.", title=COMMAND_TITLE)
            return None
        config.stud_family_name, config.stud_type_name = parse_family_type_label(
            str(stud_sel)
        )

        if self.cb_plate_type.SelectedItem:
            family, type_name = parse_family_type_label(str(self.cb_plate_type.SelectedItem))
            config.bottom_plate_family_name = family
            config.bottom_plate_type_name = type_name
            config.top_plate_family_name = family
            config.top_plate_type_name = type_name

        if self.cb_header_type.SelectedItem:
            family, type_name = parse_family_type_label(str(self.cb_header_type.SelectedItem))
            config.header_family_name = family
            config.header_type_name = type_name

        config.top_plate_count = 1 if self.rb_single_top.IsChecked else 2
        config.include_mid_plates = bool(self.chk_mid_plates.IsChecked)
        try:
            config.mid_plate_interval_ft = float(self.tb_mid_plate_height.Text)
        except Exception:
            forms.alert("Invalid mid-plate interval.", title=COMMAND_TITLE)
            return None
        config.include_king_studs = bool(self.chk_king_studs.IsChecked)
        config.include_jack_studs = bool(self.chk_jack_studs.IsChecked)
        config.include_cripple_studs = bool(self.chk_cripple_studs.IsChecked)
        config.wall_base_mode = (
            WALL_BASE_MODE_SUPPORT_TOP
            if bool(self.chk_support_top.IsChecked)
            else WALL_BASE_MODE_WALL
        )
        config.clean_existing_wall_members = bool(self.chk_clean_existing.IsChecked)
        config.track_members = True
        return config

    def _save_last(self):
        data = {
            "stud": str(self.cb_stud_type.SelectedItem or ""),
            "plate": str(self.cb_plate_type.SelectedItem or ""),
            "header": str(self.cb_header_type.SelectedItem or ""),
            "spacing_16": bool(self.rb_16oc.IsChecked),
            "spacing_24": bool(self.rb_24oc.IsChecked),
            "spacing_custom": bool(self.rb_custom.IsChecked),
            "custom_spacing": self.tb_custom_spacing.Text,
            "single_top": bool(self.rb_single_top.IsChecked),
            "mid_plates": bool(self.chk_mid_plates.IsChecked),
            "mid_interval": self.tb_mid_plate_height.Text,
            "king": bool(self.chk_king_studs.IsChecked),
            "jack": bool(self.chk_jack_studs.IsChecked),
            "cripple": bool(self.chk_cripple_studs.IsChecked),
            "support_top": bool(self.chk_support_top.IsChecked),
            "clean_existing": bool(self.chk_clean_existing.IsChecked),
        }
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
        self._select_if_present(self.cb_plate_type, data.get("plate"))
        self._select_if_present(self.cb_header_type, data.get("header"))

        if data.get("spacing_24"):
            self.rb_24oc.IsChecked = True
        elif data.get("spacing_custom"):
            self.rb_custom.IsChecked = True
            self.tb_custom_spacing.Text = str(data.get("custom_spacing", "16"))
        else:
            self.rb_16oc.IsChecked = True

        self.rb_single_top.IsChecked = bool(data.get("single_top", False))
        self.rb_double_top.IsChecked = not self.rb_single_top.IsChecked
        self.chk_mid_plates.IsChecked = bool(data.get("mid_plates", True))
        self.tb_mid_plate_height.Text = str(data.get("mid_interval", "8"))
        self.tb_mid_plate_height.IsEnabled = bool(self.chk_mid_plates.IsChecked)
        self.chk_king_studs.IsChecked = bool(data.get("king", True))
        self.chk_jack_studs.IsChecked = bool(data.get("jack", True))
        self.chk_cripple_studs.IsChecked = bool(data.get("cripple", True))
        self.chk_support_top.IsChecked = bool(data.get("support_top", False))
        self.chk_clean_existing.IsChecked = True

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


class _WallSupportFilter(ISelectionFilter):
    def AllowElement(self, element):
        if isinstance(element, (DB.Floor, DB.RoofBase)):
            return True
        category = getattr(element, "Category", None)
        if category is None:
            return False
        try:
            return int(category.Id.IntegerValue) == int(DB.BuiltInCategory.OST_StructuralFoundation)
        except Exception:
            return False

    def AllowReference(self, reference, point):
        return False


def _pick_support(doc):
    try:
        ref = revit.uidoc.Selection.PickObject(
            ObjectType.Element,
            _WallSupportFilter(),
            "Pick floor, roof, or foundation for wall base",
        )
    except Exception:
        return None
    if ref is None:
        return None
    return doc.GetElement(ref.ElementId)


def _support_top_elevation(element):
    if element is None:
        return None
    try:
        bbox = element.get_BoundingBox(None)
        if bbox is not None:
            return bbox.Max.Z
    except Exception:
        pass
    return None


def _delete_existing_wall_members(doc, walls, include_legacy):
    wall_ids = set()
    for wall in walls or []:
        wall_id = _element_id_text(getattr(wall, "Id", None))
        if wall_id:
            wall_ids.add(wall_id)
    if not wall_ids:
        return {"wall_v4": 0, "wall_v2": 0, "wall": 0}

    allowed_kinds = set(["wall_v4", "wall_v2"])
    if include_legacy:
        allowed_kinds.add("wall")

    delete_items = []
    for category in (
        DB.BuiltInCategory.OST_StructuralFraming,
        DB.BuiltInCategory.OST_StructuralColumns,
    ):
        try:
            collector = (
                DB.FilteredElementCollector(doc)
                .OfCategory(category)
                .WhereElementIsNotElementType()
            )
        except Exception:
            continue
        for element in collector:
            tracking = get_tracking_data(element)
            if not tracking:
                continue
            kind = tracking.get("kind")
            if kind not in allowed_kinds:
                continue
            if tracking.get("host") in wall_ids:
                delete_items.append((element.Id, kind))

    deleted = {"wall_v4": 0, "wall_v2": 0, "wall": 0}
    for element_id, kind in delete_items:
        try:
            doc.Delete(element_id)
            deleted[kind] = deleted.get(kind, 0) + 1
        except Exception:
            pass
    return deleted


def _audit_line(audit, placed_count):
    wall_id = audit.get("wall_id", "?")
    source_side = audit.get("source_side") or ""
    target_offset = float(audit.get("target_offset_from_face") or 0.0)
    first_rejection = audit.get("first_rejection") or ""
    if first_rejection:
        first_rejection = str(first_rejection).replace("|", "/").replace("\n", " ")[:80]
    return (
        "| {0} | {1:.3f} | {2:.3f} | {3} | {4} | {5} | {6} | {7} | {8} | {9} | {10} | {11} | {12:.3f} | {13} |\n"
        .format(
            wall_id,
            float(audit.get("location_length") or 0.0),
            float(audit.get("face_length") or 0.0),
            audit.get("wall_solid_count", 0),
            audit.get("face_loop_count", 0),
            audit.get("merged_opening_count", 0),
            audit.get("perimeter_segment_count", 0),
            audit.get("candidate_count", 0),
            audit.get("validated_count", 0),
            audit.get("rejected_count", 0),
            placed_count,
            source_side,
            target_offset,
            first_rejection,
        )
    )


def _element_id_text(element_id):
    if element_id is None:
        return None
    value = getattr(element_id, "IntegerValue", getattr(element_id, "Value", None))
    if value is None:
        return None
    return str(value)


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

    selected = revit.get_selection().elements
    walls = [element for element in selected if isinstance(element, DB.Wall)]
    if not walls:
        try:
            refs = revit.uidoc.Selection.PickObjects(
                ObjectType.Element,
                _WallFilter(),
                "Select walls for wall framing",
            )
            walls = [doc.GetElement(ref.ElementId) for ref in refs]
        except Exception:
            return
    if not walls:
        return

    dialog = WallFraming20Dialog(column_types, framing_types, len(walls))
    dialog.ShowDialog()
    config = dialog.result_config
    if config is None:
        return

    if config.wall_base_mode == WALL_BASE_MODE_SUPPORT_TOP:
        support = _pick_support(doc)
        if support is None:
            forms.alert("Wall base selection was cancelled.", title=COMMAND_TITLE)
            return
        config.wall_base_override_z = _support_top_elevation(support)

    engine = WallCavityFramingV4Engine(doc, config)
    wall_count = 0
    member_count = 0
    skipped = 0
    deleted_counts = {"wall_v4": 0, "wall_v2": 0, "wall": 0}
    audit_rows = []

    with revit.Transaction("WF: Wall Framing"):
        deleted_counts = _delete_existing_wall_members(
            doc,
            walls,
            bool(getattr(config, "clean_existing_wall_members", True)),
        )
        for wall in walls:
            members, host = engine.calculate_members(wall)
            if host is None:
                skipped += 1
                continue
            placed = engine.place_members(members, host)
            wall_count += 1
            member_count += len(placed)
            audit_rows.append((getattr(host, "audit", {}), len(placed)))

    audit_table = (
        "\n### Source Geometry Audit\n"
        "| Wall Id | Location Len | Face Len | Solids | Face Loops | Openings | Perimeter Segs | Candidates | Validated | Rejected | Placed | Source Side | Target Offset | First Rejection |\n"
        "| --- | ---: | ---: | ---: | ---: | ---: | ---: | ---: | ---: | ---: | ---: | --- | ---: | --- |\n"
    )
    for audit, placed_count in audit_rows:
        audit_table += _audit_line(audit, placed_count)

    deleted_total = (
        deleted_counts.get("wall_v4", 0)
        + deleted_counts.get("wall_v2", 0)
        + deleted_counts.get("wall", 0)
    )

    output.print_md(
        "## Wall Framing Complete\n"
        "- **Engine:** {0}\n"
        "- **Walls framed:** {1}\n"
        "- **Walls skipped:** {2}\n"
        "- **Previous members replaced:** {3}\n"
        "- **Members placed:** {4}\n"
        "{5}".format(
            ENGINE_NAME,
            wall_count,
            skipped,
            deleted_total,
            member_count,
            audit_table,
        )
    )


if __name__ == "__main__":
    main()
