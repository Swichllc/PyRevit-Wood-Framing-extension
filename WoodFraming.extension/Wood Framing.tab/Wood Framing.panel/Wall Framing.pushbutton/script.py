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
    WALL_JOIN_CORNER_INSULATED,
    WALL_JOIN_T_ASSEMBLY,
    WALL_BASE_MODE_SUPPORT_TOP,
    WALL_BASE_MODE_WALL,
)
from wf_families import (
    get_available_types_flat,
    get_column_types_flat,
    parse_family_type_label,
)
from wf_tracking import get_tracking_data
from wf_wall_framing_v4 import ENGINE_NAME, STUD_THICKNESS, WallCavityFramingV4Engine
from wf_wall_join_cleanup import (
    JOIN_KIND_CORNER,
    JOIN_KIND_T,
    STYLE_CORNER_CAVITY,
    STYLE_CORNER_INSULATED,
    STYLE_T_ASSEMBLY,
    STYLE_T_BLOCKING_NAILER,
    build_wall_join_assembly_plans,
)


logger = script.get_logger()
output = script.get_output()
COMMAND_TITLE = "Wall Framing"
_XAML = os.path.join(os.path.dirname(__file__), "FrameWall20Config.xaml")
_CFG_PATH = os.path.join(
    os.environ.get("APPDATA", ""),
    "pyRevit",
    "WoodFraming_Wall20LastConfig.json",
)

CORNER_STYLE_LABELS = [
    ("Insulated corner - two studs", STYLE_CORNER_INSULATED),
    ("Cavity corner - two studs plus blocking", STYLE_CORNER_CAVITY),
]
T_STYLE_LABELS = [
    ("T assembly", STYLE_T_ASSEMBLY),
    ("Blocking and 1x nailer", STYLE_T_BLOCKING_NAILER),
]
JOIN_CONFLICT_MEMBER_TYPES = set(["SIDE_STUD", "STUD"])


class WallFraming20Dialog(WPFWindow):
    def __init__(self, column_types, framing_types, wall_count):
        WPFWindow.__init__(self, _XAML)
        self.result_config = None
        self._column_types = column_types
        self._framing_types = framing_types

        self.cb_stud_type.ItemsSource = column_types
        self.cb_plate_type.ItemsSource = framing_types
        self.cb_header_type.ItemsSource = framing_types
        self.cb_corner_assembly.ItemsSource = [
            label for label, _key in CORNER_STYLE_LABELS
        ]
        self.cb_t_assembly.ItemsSource = [
            label for label, _key in T_STYLE_LABELS
        ]

        if column_types:
            self.cb_stud_type.SelectedIndex = 0
        if framing_types:
            self.cb_plate_type.SelectedIndex = 0
            self.cb_header_type.SelectedIndex = 0
        if CORNER_STYLE_LABELS:
            self.cb_corner_assembly.SelectedIndex = 0
        if T_STYLE_LABELS:
            self.cb_t_assembly.SelectedIndex = 0

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
        config.wall_join_corner_style = _selected_style_key(
            self.cb_corner_assembly,
            CORNER_STYLE_LABELS,
            WALL_JOIN_CORNER_INSULATED,
        )
        config.wall_join_t_style = _selected_style_key(
            self.cb_t_assembly,
            T_STYLE_LABELS,
            WALL_JOIN_T_ASSEMBLY,
        )
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
            "corner_assembly": str(self.cb_corner_assembly.SelectedItem or ""),
            "t_assembly": str(self.cb_t_assembly.SelectedItem or ""),
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
        self._select_if_present(self.cb_corner_assembly, data.get("corner_assembly"))
        self._select_if_present(self.cb_t_assembly, data.get("t_assembly"))

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


def _selected_style_key(combo, style_pairs, fallback):
    selected = str(getattr(combo, "SelectedItem", "") or "")
    for label, key in style_pairs:
        if selected == label:
            return key
    return fallback


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
        return {"wall_v4": 0, "wall_v2": 0, "wall": 0, "wall_join": 0}

    allowed_kinds = set(["wall_v4", "wall_v2"])
    if len(wall_ids) >= 2:
        allowed_kinds.add("wall_join")
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

    deleted = {"wall_v4": 0, "wall_v2": 0, "wall": 0, "wall_join": 0}
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


def _audit_text(value, limit):
    text = str(value or "")
    text = text.replace("|", "/").replace("\n", " ")
    if len(text) > limit:
        return text[:max(0, limit - 3)] + "..."
    return text


def _side_stud_attempt_summary(audit):
    attempts = audit.get("side_stud_attempts") or []
    parts = []
    for attempt in attempts[:6]:
        try:
            raw_d = float(attempt.get("raw_d") or 0.0)
            stud_d = float(attempt.get("stud_d") or 0.0)
        except Exception:
            raw_d = 0.0
            stud_d = 0.0
        result = attempt.get("result") or "?"
        label = "{0:.3f}->{1:.3f} {2}".format(raw_d, stud_d, result)
        reason = attempt.get("reason") or ""
        if reason:
            label = "{0} ({1})".format(label, reason)
        parts.append(label)
    if len(attempts) > 6:
        parts.append("+{0} more".format(len(attempts) - 6))
    return _audit_text("; ".join(parts), 120)


def _side_stud_audit_line(audit):
    wall_id = audit.get("wall_id", "?")
    return (
        "| {0} | {1} | {2} | {3} | {4} | {5} |\n"
        .format(
            wall_id,
            audit.get("side_segment_count", 0),
            audit.get("side_stud_attempted_count", 0),
            audit.get("side_stud_placed_count", 0),
            audit.get("side_stud_rejected_count", 0),
            _side_stud_attempt_summary(audit),
        )
    )


def _element_id_text(element_id):
    if element_id is None:
        return None
    value = getattr(element_id, "IntegerValue", getattr(element_id, "Value", None))
    if value is None:
        return None
    return str(value)


class _WallFrameStage(object):
    def __init__(self, wall, host, members, occupied):
        self.wall = wall
        self.host = host
        self.members = members
        self.occupied = occupied


def _calculate_wall_stage(engine, wall):
    host = engine._analyze_wall(wall)
    if host is None:
        return None
    occupied = set()
    members = []
    members.extend(engine._wall_shape_members(host, occupied))
    members.extend(engine._opening_members(host, occupied))
    return _WallFrameStage(wall, host, members, occupied)


def _finish_wall_stage(engine, stage, join_members):
    _mark_join_members_occupied(stage.host, stage.occupied, join_members)
    stage.members, skipped = _filter_join_conflicts(
        stage.host,
        stage.members,
        join_members,
    )
    stage.members.extend(engine._infill_members(stage.host, stage.occupied))
    return skipped


def _join_members_by_host_id(join_plans):
    result = {}
    for plan in join_plans or []:
        for member in getattr(plan, "members", []) or []:
            host_id = _element_id_text(getattr(member, "host_id", None))
            if host_id is None:
                continue
            result.setdefault(host_id, []).append(member)
    return result


def _stage_host_map(stages):
    result = {}
    for stage in stages or []:
        host = getattr(stage, "host", None)
        host_id = _element_id_text(getattr(host, "element_id", None))
        if host_id is not None:
            result[host_id] = host
    return result


def _mark_join_members_occupied(host, occupied, join_members):
    for member in join_members or []:
        if not getattr(member, "is_column", False):
            continue
        d = _member_distance_on_host(host, member)
        if d is None:
            continue
        if d < -STUD_THICKNESS or d > host.length + STUD_THICKNESS:
            continue
        occupied.add(round(max(0.0, min(host.length, d)), 4))


def _filter_join_conflicts(host, members, join_members):
    join_distances = []
    for join_member in join_members or []:
        if not getattr(join_member, "is_column", False):
            continue
        d = _member_distance_on_host(host, join_member)
        if d is not None:
            join_distances.append(d)
    if not join_distances:
        return members, 0

    kept = []
    skipped = 0
    for member in members:
        member_type = str(getattr(member, "member_type", "") or "").upper()
        if (member_type not in JOIN_CONFLICT_MEMBER_TYPES
                or not getattr(member, "is_column", False)):
            kept.append(member)
            continue
        d = _member_distance_on_host(host, member)
        if d is None:
            kept.append(member)
            continue
        if _near_any_distance(d, join_distances, STUD_THICKNESS * 0.75):
            skipped += 1
            continue
        kept.append(member)
    return kept, skipped


def _member_distance_on_host(host, member):
    point = getattr(member, "start_point", None)
    if point is None:
        return None
    try:
        vec = point - host.start_point
        return vec.DotProduct(host.direction)
    except Exception:
        return None


def _near_any_distance(value, distances, tolerance):
    for distance in distances:
        if abs(value - distance) <= tolerance:
            return True
    return False


def _place_join_plans(engine, join_plans):
    placed_count = 0
    for plan in join_plans or []:
        for host in getattr(plan, "hosts", []) or []:
            host_id = _element_id_text(getattr(host, "element_id", None))
            host_members = [
                member for member in getattr(plan, "members", []) or []
                if _element_id_text(getattr(member, "host_id", None)) == host_id
            ]
            if not host_members:
                continue
            placed_count += len(engine.place_members(host_members, host))
    return placed_count


def _join_audit_table(join_plans):
    table = (
        "\n### Join Assembly Audit\n"
        "| Join | Assembly | Members |\n"
        "| --- | --- | ---: |\n"
    )
    if not join_plans:
        return table + "| None detected | | 0 |\n"
    for plan in join_plans:
        table += (
            "| {0} | {1} | {2} |\n"
            .format(
                _join_label(getattr(plan, "join_kind", None)),
                _style_label(getattr(plan, "style_key", None)),
                len(getattr(plan, "members", []) or []),
            )
        )
    return table


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
    deleted_counts = {"wall_v4": 0, "wall_v2": 0, "wall": 0, "wall_join": 0}
    audit_rows = []
    join_plans = []
    join_requested = 0
    join_placed = 0
    join_conflict_skipped = 0

    with revit.Transaction("WF: Wall Framing"):
        deleted_counts = _delete_existing_wall_members(
            doc,
            walls,
            bool(getattr(config, "clean_existing_wall_members", True)),
        )

        stages = []
        for wall in walls:
            stage = _calculate_wall_stage(engine, wall)
            if stage is None:
                skipped += 1
                continue
            stages.append(stage)

        join_plans = build_wall_join_assembly_plans(
            doc,
            [stage.wall for stage in stages],
            config,
            getattr(config, "wall_join_corner_style", STYLE_CORNER_INSULATED),
            getattr(config, "wall_join_t_style", STYLE_T_ASSEMBLY),
            _stage_host_map(stages),
        )
        join_requested = sum(
            len(getattr(plan, "members", []) or [])
            for plan in join_plans
        )
        join_members_by_host = _join_members_by_host_id(join_plans)

        for stage in stages:
            host_id = _element_id_text(getattr(stage.host, "element_id", None))
            host_join_members = join_members_by_host.get(host_id, [])
            join_conflict_skipped += _finish_wall_stage(
                engine,
                stage,
                host_join_members,
            )
            placed = engine.place_members(stage.members, stage.host)
            wall_count += 1
            member_count += len(placed)
            audit_rows.append((getattr(stage.host, "audit", {}), len(placed)))

        join_placed = _place_join_plans(engine, join_plans)
        member_count += join_placed

    audit_table = (
        "\n### Source Geometry Audit\n"
        "| Wall Id | Location Len | Face Len | Solids | Face Loops | Openings | Perimeter Segs | Candidates | Validated | Rejected | Placed | Source Side | Target Offset | First Rejection |\n"
        "| --- | ---: | ---: | ---: | ---: | ---: | ---: | ---: | ---: | ---: | ---: | --- | ---: | --- |\n"
    )
    for audit, placed_count in audit_rows:
        audit_table += _audit_line(audit, placed_count)

    side_stud_table = (
        "\n### Side Stud Audit\n"
        "| Wall Id | Side Segs | Attempted | Placed | Rejected | Attempts |\n"
        "| --- | ---: | ---: | ---: | ---: | --- |\n"
    )
    for audit, _placed_count in audit_rows:
        side_stud_table += _side_stud_audit_line(audit)

    join_table = _join_audit_table(join_plans)

    deleted_total = (
        deleted_counts.get("wall_v4", 0)
        + deleted_counts.get("wall_v2", 0)
        + deleted_counts.get("wall", 0)
        + deleted_counts.get("wall_join", 0)
    )

    output.print_md(
        "## Wall Framing Complete\n"
        "- **Engine:** {0}\n"
        "- **Walls framed:** {1}\n"
        "- **Walls skipped:** {2}\n"
        "- **Previous members replaced:** {3}\n"
        "- **Members placed:** {4}\n"
        "- **Join assemblies detected:** {5}\n"
        "- **Join assembly members requested:** {6}\n"
        "- **Join assembly members placed:** {7}\n"
        "- **Base side studs skipped for join assemblies:** {8}\n"
        "{9}"
        "{10}"
        "{11}".format(
            ENGINE_NAME,
            wall_count,
            skipped,
            deleted_total,
            member_count,
            len(join_plans),
            join_requested,
            join_placed,
            join_conflict_skipped,
            audit_table,
            side_stud_table,
            join_table,
        )
    )


if __name__ == "__main__":
    main()
