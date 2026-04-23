# -*- coding: utf-8 -*-
"""Frame Wall — wood framing for selected Revit walls.

Studs (vertical)   -> Structural Column families (OST_StructuralColumns)
Plates / Headers    -> Structural Framing families (OST_StructuralFraming)

Construction sequence per wall:
  A. Bottom plate(s)
  B. Mid plates (fire blocking every 8 ft) + top plates (follow wall profile)
  C. Opening framing: king studs, jack/trimmer studs
  D. Opening headers (doubled/tripled) + sill plates + cripple studs
  E. Corner / end studs
  F. Infill studs at OC spacing (split at mid-plate tiers)
  G. Blocking / bridging at mid-height for tall stud runs
"""

import os
import sys
import math

from pyrevit import revit, DB, script, forms
from Autodesk.Revit.UI.Selection import ObjectType, ISelectionFilter

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
    LAYER_MODE_CORE_CENTER,
    LAYER_MODE_STRUCTURAL,
    LAYER_MODE_THICKEST,
    SPACING_16OC,
    SPACING_24OC,
    WALL_BASE_MODE_WALL,
    WALL_BASE_MODE_SUPPORT_TOP,
)
from wf_families import (
    get_available_types_flat,
    get_column_types_flat,
    parse_family_type_label,
    find_family_symbol,
    activate_symbol,
)
from wf_host import analyze_wall_host
from wf_geometry import safe_wall_normal
from wf_wall_joins import build_wall_join_plan

logger = script.get_logger()
output = script.get_output()

# ---------------------------------------------------------------------------
# Constants
# ---------------------------------------------------------------------------
PLATE_THICKNESS = 1.5 / 12.0   # 1.5 in -> feet
STUD_THICKNESS  = 1.5 / 12.0
MIN_LENGTH      = 1.0 / 12.0   # 1 in minimum member
MID_PLATE_INTERVAL = 8.0       # fire blocking every 8 ft
BLOCKING_MAX_HEIGHT = 8.0      # add mid-height blocking if stud > 8 ft
PLATE_ROT = -math.pi / 2.0     # plates lay flat

# Config persistence (remembers last dialog settings across sessions)
_CFG_PATH = os.path.join(
    os.environ.get("APPDATA", ""),
    "pyRevit", "WoodFraming_LastConfig.json",
)


# =========================================================================
#  Placement helpers
# =========================================================================

def _find_symbol(doc, fam_name, type_name, category):
    sym = find_family_symbol(doc, fam_name, type_name, category)
    if sym is not None:
        activate_symbol(doc, sym)
    return sym


def place_beam(doc, start_xyz, end_xyz, symbol, level, rotation=0.0):
    """Place Structural Framing along a line."""
    if start_xyz.DistanceTo(end_xyz) < MIN_LENGTH:
        return None
    line = DB.Line.CreateBound(start_xyz, end_xyz)
    try:
        inst = doc.Create.NewFamilyInstance(
            line, symbol, level, DB.Structure.StructuralType.Beam
        )
    except Exception:
        return None
    if inst is None:
        return None
    _set_param_int(inst, DB.BuiltInParameter.YZ_JUSTIFICATION, 0)
    _set_param_int(inst, DB.BuiltInParameter.Y_JUSTIFICATION, 2)
    _set_param_int(inst, DB.BuiltInParameter.Z_JUSTIFICATION, 2)
    for pid in (
        DB.BuiltInParameter.Y_OFFSET_VALUE,
        DB.BuiltInParameter.Z_OFFSET_VALUE,
        DB.BuiltInParameter.START_Y_OFFSET_VALUE,
        DB.BuiltInParameter.END_Y_OFFSET_VALUE,
        DB.BuiltInParameter.START_Z_OFFSET_VALUE,
        DB.BuiltInParameter.END_Z_OFFSET_VALUE,
    ):
        _set_param_double(inst, pid, 0.0)
    _set_param_double(inst, DB.BuiltInParameter.STRUCTURAL_BEND_DIR_ANGLE, rotation)
    _tag(inst)
    return inst


def place_column(doc, base_xyz, height, symbol, level, wall_angle):
    """Place Structural Column at a point."""
    if height < MIN_LENGTH:
        return None
    try:
        inst = doc.Create.NewFamilyInstance(
            base_xyz, symbol, level, DB.Structure.StructuralType.Column
        )
    except Exception:
        return None
    if inst is None:
        return None
    _set_param_int(inst, DB.BuiltInParameter.YZ_JUSTIFICATION, 0)
    _set_param_int(inst, DB.BuiltInParameter.Y_JUSTIFICATION, 2)
    _set_param_int(inst, DB.BuiltInParameter.Z_JUSTIFICATION, 2)
    for pid in (
        DB.BuiltInParameter.Y_OFFSET_VALUE,
        DB.BuiltInParameter.Z_OFFSET_VALUE,
        DB.BuiltInParameter.START_Y_OFFSET_VALUE,
        DB.BuiltInParameter.END_Y_OFFSET_VALUE,
        DB.BuiltInParameter.START_Z_OFFSET_VALUE,
        DB.BuiltInParameter.END_Z_OFFSET_VALUE,
    ):
        _set_param_double(inst, pid, 0.0)
    for bip in (DB.BuiltInParameter.FAMILY_TOP_LEVEL_PARAM,
                DB.BuiltInParameter.FAMILY_BASE_LEVEL_PARAM):
        try:
            p = inst.get_Parameter(bip)
            if p and not p.IsReadOnly:
                p.Set(level.Id)
        except Exception:
            pass
    lev_elev = level.Elevation
    base_off = base_xyz.Z - lev_elev
    _set_param_double(inst, DB.BuiltInParameter.FAMILY_BASE_LEVEL_OFFSET_PARAM, base_off)
    _set_param_double(inst, DB.BuiltInParameter.FAMILY_TOP_LEVEL_OFFSET_PARAM, base_off + height)
    if wall_angle != 0.0:
        try:
            axis = DB.Line.CreateBound(base_xyz, base_xyz + DB.XYZ.BasisZ)
            DB.ElementTransformUtils.RotateElement(doc, inst.Id, axis, wall_angle)
        except Exception:
            pass
    _tag(inst)
    return inst


def _set_param_int(inst, bip, value):
    try:
        p = inst.get_Parameter(bip)
        if p and not p.IsReadOnly:
            p.Set(value)
    except Exception:
        pass


def _set_param_double(inst, bip, value):
    try:
        p = inst.get_Parameter(bip)
        if p and not p.IsReadOnly:
            p.Set(value)
    except Exception:
        pass


def _tag(inst):
    try:
        p = inst.get_Parameter(DB.BuiltInParameter.ALL_MODEL_INSTANCE_COMMENTS)
        if p and not p.IsReadOnly:
            p.Set("WF_Generated")
    except Exception:
        pass


def _get_depth(symbol):
    if symbol is None:
        return PLATE_THICKNESS
    try:
        p = symbol.LookupParameter("d")
        if p and p.HasValue:
            v = p.AsDouble()
            if v > 0:
                return v
    except Exception:
        pass
    return PLATE_THICKNESS


# =========================================================================
#  Wall framing engine
# =========================================================================

class WallFramer(object):
    def __init__(self, doc, config):
        self.doc = doc
        self.config = config
        self.placed = []

    def frame_wall(self, wall):
        loc_curve = wall.Location
        if loc_curve is None:
            return False
        curve = loc_curve.Curve
        if not isinstance(curve, DB.Line):
            return False

        length = curve.Length
        p0 = curve.GetEndPoint(0)
        p1 = curve.GetEndPoint(1)
        direction = (p1 - p0).Normalize()
        normal = safe_wall_normal(wall, direction)
        if normal is None:
            return False
        wall_angle = math.atan2(direction.Y, direction.X)

        host_info = None
        target_offset = 0.0
        try:
            host_info = analyze_wall_host(self.doc, wall, self.config)
            if host_info is not None:
                target_offset = host_info.target_layer_offset
        except Exception:
            pass
        if host_info is not None:
            direction = host_info.direction
            normal = host_info.normal
            length = host_info.length
            wall_angle = host_info.wall_info.angle
            level = self.doc.GetElement(host_info.level_id)
            if level is None:
                return False
            base_z = host_info.base_elevation
            p0 = host_info.point_at(0.0, 0.0)
            p1 = host_info.point_at(length, 0.0)
            # Read the wall's own profile — sketch for edited shapes,
            # bounding-box for standard rectangular walls.
            top_profile = self._wall_top_profile(
                wall, p0, direction, base_z, length)
            openings = self._opening_dicts_from_host_info(host_info)
        else:
            p0 = p0 + normal * target_offset
            p1 = p1 + normal * target_offset

            level = self.doc.GetElement(wall.LevelId)
            if level is None:
                return False
            base_z = level.Elevation
            p_off = wall.get_Parameter(DB.BuiltInParameter.WALL_BASE_OFFSET)
            if p_off and p_off.HasValue:
                base_z += p_off.AsDouble()
            p0 = DB.XYZ(p0.X, p0.Y, base_z)
            p1 = DB.XYZ(p1.X, p1.Y, base_z)

            h_param = wall.get_Parameter(DB.BuiltInParameter.WALL_USER_HEIGHT_PARAM)
            wall_height = h_param.AsDouble() if (h_param and h_param.HasValue) else 8.0

            top_profile = self._wall_top_profile(
                wall, p0, direction, base_z, length)
            openings = self._get_openings(wall, p0, direction, base_z)

        join_plan = None
        if host_info is not None:
            try:
                join_plan = build_wall_join_plan(
                    self.doc,
                    host_info,
                    self.config,
                    STUD_THICKNESS,
                )
            except Exception:
                join_plan = None

        framing_cat = DB.BuiltInCategory.OST_StructuralFraming
        column_cat  = DB.BuiltInCategory.OST_StructuralColumns

        plate_sym = _find_symbol(
            self.doc,
            self.config.bottom_plate_family_name or self.config.stud_family_name,
            self.config.bottom_plate_type_name or self.config.stud_type_name,
            framing_cat,
        )
        header_sym = _find_symbol(
            self.doc,
            self.config.header_family_name or self.config.stud_family_name,
            self.config.header_type_name or self.config.stud_type_name,
            framing_cat,
        )
        stud_sym = _find_symbol(
            self.doc,
            self.config.stud_family_name,
            self.config.stud_type_name,
            column_cat,
        )

        if stud_sym is None:
            logger.warning("Stud column symbol not found")
            return False
        if plate_sym is None:
            logger.warning("Plate framing symbol not found")
            return False

        stud_bottom = base_z + PLATE_THICKNESS * self.config.bottom_plate_count

        door_gaps = [(op["left"], op["right"]) for op in openings if not op["is_window"]]
        plate_segs = self._split_segments(0.0, length, door_gaps)

        occupied = set()
        spacing = self.config.stud_spacing_ft
        # Track placed beams/columns for attachment pass
        _plate_beams = []   # (inst, seg_start_d, seg_end_d, is_angled)
        _stud_columns = []  # (inst, d_position)

        # ==== A. BOTTOM PLATES ====
        for i in range(self.config.bottom_plate_count):
            pz = base_z + i * PLATE_THICKNESS + PLATE_THICKNESS / 2.0
            for seg_s, seg_e in plate_segs:
                s = p0 + direction * seg_s + DB.XYZ(0, 0, pz - base_z)
                e = p0 + direction * seg_e + DB.XYZ(0, 0, pz - base_z)
                self._place(place_beam(self.doc, s, e, plate_sym, level, PLATE_ROT))

        # ==== B. MID PLATES + TOP PLATES ====
        min_top_z = min(z for _, z in top_profile)
        min_stud_top = min_top_z - PLATE_THICKNESS * self.config.top_plate_count
        total_stud_height = min_stud_top - stud_bottom
        mid_plate_zs = []
        if bool(getattr(self.config, "include_mid_plates", True)):
            interval = float(getattr(self.config, "mid_plate_interval_ft", MID_PLATE_INTERVAL))
            # Mid-plate positions are measured upward from bottom-plate stack top.
            if interval > 0.0 and total_stud_height > interval + PLATE_THICKNESS * 1.5:
                tier = 1
                while True:
                    mz = stud_bottom + tier * interval
                    if mz < min_stud_top - PLATE_THICKNESS:
                        mid_plate_zs.append(mz)
                        tier += 1
                        continue
                    break

        # Mid plates split at ALL openings (doors + windows), not just doors
        all_opening_gaps = [(op["left"], op["right"]) for op in openings]
        mid_plate_segs = self._split_segments(0.0, length, all_opening_gaps)

        for mz in mid_plate_zs:
            mid_center = mz + PLATE_THICKNESS / 2.0
            for seg_s, seg_e in mid_plate_segs:
                s = p0 + direction * seg_s + DB.XYZ(0, 0, mid_center - base_z)
                e = p0 + direction * seg_e + DB.XYZ(0, 0, mid_center - base_z)
                self._place(place_beam(self.doc, s, e, plate_sym, level, PLATE_ROT))

        # Top plates — run full wall length (no trim at corners)
        key_pts = self._simplify_profile(top_profile)
        if key_pts:
            z_vals = [z for _, z in key_pts]
            wall_is_sloped = (max(z_vals) - min(z_vals)) > (1.0 / 24.0)
        else:
            wall_is_sloped = False
        for si in range(len(key_pts) - 1):
            d0, z0 = key_pts[si]
            d1, z1 = key_pts[si + 1]
            seg_s = d0
            seg_e = d1
            if seg_e - seg_s < MIN_LENGTH:
                continue

            if abs(d1 - d0) > 1e-9:
                t0 = (seg_s - d0) / (d1 - d0)
                t1 = (seg_e - d0) / (d1 - d0)
                z_start = z0 + (z1 - z0) * t0
                z_end = z0 + (z1 - z0) * t1
            else:
                z_start = z0
                z_end = z1

            for i in range(self.config.top_plate_count):
                off = (self.config.top_plate_count - i - 1) * PLATE_THICKNESS
                off += PLATE_THICKNESS / 2.0
                s = p0 + direction * seg_s + DB.XYZ(0, 0, z_start - off - base_z)
                e = p0 + direction * seg_e + DB.XYZ(0, 0, z_end - off - base_z)
                inst = place_beam(self.doc, s, e, plate_sym, level, PLATE_ROT)
                self._place(inst)
                # Track the lowest top-plate run on sloped walls for stud attachment.
                # A wall can have short flat segments even when the overall top is angled.
                if inst and i == 0 and wall_is_sloped:
                    seg_is_angled = abs(z_end - z_start) > (1.0 / 96.0)
                    _plate_beams.append((inst, seg_s, seg_e, seg_is_angled))

        # ==== C. KING STUDS & JACK STUDS ====
        # Jack stud defines the RO — inner face at RO edge.
        # King stud is full-height, nailed to jack on the wall side.
        for op in openings:
            left_e  = op["left"]   # rough opening left edge
            right_e = op["right"]  # rough opening right edge
            head_z  = op["head_z"]

            # Jack studs: inner face at RO edge → center half-stud outside RO
            jack_left  = left_e  - STUD_THICKNESS / 2.0
            jack_right = right_e + STUD_THICKNESS / 2.0
            # King studs: outside the jack → center 1.5 studs outside RO
            king_left  = left_e  - STUD_THICKNESS * 1.5
            king_right = right_e + STUD_THICKNESS * 1.5

            if self.config.include_king_studs:
                for kd in (king_left, king_right):
                    kd_c = max(0.0, min(length, kd))
                    if self._near(kd_c, occupied):
                        continue
                    # Split king studs at mid plates
                    local_top = self._top_of_stud_at(top_profile, kd_c)
                    tier_bounds = [stud_bottom] + list(mid_plate_zs) + [local_top]
                    last_tier = len(tier_bounds) - 2
                    for ti in range(len(tier_bounds) - 1):
                        t_bot = tier_bounds[ti]
                        t_top = tier_bounds[ti + 1]
                        if ti > 0:
                            t_bot += PLATE_THICKNESS
                        t_h = t_top - t_bot
                        if t_h >= MIN_LENGTH:
                            sl = _stud_columns if ti == last_tier else None
                            self._place_stud(p0, direction, kd_c, t_bot,
                                             t_h, stud_sym, level, wall_angle, base_z,
                                             sl)
                    occupied.add(round(kd_c, 4))

            if self.config.include_jack_studs and head_z > stud_bottom:
                jack_h = head_z - stud_bottom
                for jd in (jack_left, jack_right):
                    jd_c = max(0.0, min(length, jd))
                    if self._near(jd_c, occupied):
                        continue
                    self._place_stud(p0, direction, jd_c, stud_bottom,
                                     jack_h, stud_sym, level, wall_angle, base_z)
                    occupied.add(round(jd_c, 4))

        # ==== D. HEADERS + SILLS + CRIPPLES ====
        for op in openings:
            left_e  = op["left"]
            right_e = op["right"]
            head_z  = op["head_z"]
            sill_z  = op["sill_z"]
            is_win  = op["is_window"]

            # Header spans between king stud inner faces.
            # With jacks: king inner face is one stud outside RO edge.
            # Without jacks: king inner face IS the RO edge.
            if self.config.include_jack_studs:
                span_s = left_e  - STUD_THICKNESS
                span_e = right_e + STUD_THICKNESS
            else:
                span_s = left_e
                span_e = right_e
            span_len = span_e - span_s

            if header_sym is not None and span_len >= MIN_LENGTH:
                h_depth = _get_depth(header_sym)
                h_count = self.config.header_count
                if span_len > 6.0:
                    h_count = max(h_count, 3)
                elif span_len < 3.0:
                    h_count = max(1, h_count)

                # Head plate (flat) at top of rough opening
                if plate_sym is not None:
                    hp_z = head_z + PLATE_THICKNESS / 2.0
                    hps = p0 + direction * span_s + DB.XYZ(0, 0, hp_z - base_z)
                    hpe = p0 + direction * span_e + DB.XYZ(0, 0, hp_z - base_z)
                    self._place(place_beam(self.doc, hps, hpe, plate_sym, level, PLATE_ROT))

                # Header member(s) on-edge above head plate — side by side
                header_base = head_z + PLATE_THICKNESS
                hz = header_base + h_depth / 2.0
                for hi in range(h_count):
                    # Offset each header in wall-normal direction (face-to-face)
                    y_off = (hi - (h_count - 1) / 2.0) * STUD_THICKNESS
                    n_shift = normal * y_off
                    hs = p0 + direction * span_s + DB.XYZ(n_shift.X, n_shift.Y, hz - base_z)
                    he = p0 + direction * span_e + DB.XYZ(n_shift.X, n_shift.Y, hz - base_z)
                    self._place(place_beam(self.doc, hs, he, header_sym, level, 0.0))

                # Top plate above header assembly
                header_top_z = header_base + h_depth
                if plate_sym is not None:
                    tp_z = header_top_z + PLATE_THICKNESS / 2.0
                    tps = p0 + direction * span_s + DB.XYZ(0, 0, tp_z - base_z)
                    tpe = p0 + direction * span_e + DB.XYZ(0, 0, tp_z - base_z)
                    self._place(place_beam(self.doc, tps, tpe, plate_sym, level, PLATE_ROT))
                    header_top_z += PLATE_THICKNESS

                # Cripple studs above header — aligned to wall OC grid
                mid_d = (left_e + right_e) / 2.0
                local_stud_top = self._top_of_stud_at(top_profile, mid_d)
                if self.config.include_cripple_studs and header_top_z < local_stud_top:
                    crip_h = local_stud_top - header_top_z
                    self._place_cripples_oc(p0, direction, left_e, right_e,
                                            header_top_z, crip_h, spacing,
                                            stud_sym, level, wall_angle, base_z,
                                            occupied)

            # Sill plate for windows
            if is_win and sill_z > stud_bottom and plate_sym:
                sill_center = sill_z - PLATE_THICKNESS / 2.0
                if right_e - left_e >= MIN_LENGTH:
                    ss = p0 + direction * left_e  + DB.XYZ(0, 0, sill_center - base_z)
                    se = p0 + direction * right_e + DB.XYZ(0, 0, sill_center - base_z)
                    self._place(place_beam(self.doc, ss, se, plate_sym, level, PLATE_ROT))

                # Cripple studs below sill — aligned to wall OC grid
                if self.config.include_cripple_studs:
                    crip_top = sill_z - PLATE_THICKNESS
                    crip_h = crip_top - stud_bottom
                    if crip_h >= MIN_LENGTH:
                        self._place_cripples_oc(p0, direction, left_e, right_e,
                                                stud_bottom, crip_h, spacing,
                                                stud_sym, level, wall_angle, base_z,
                                                occupied)

        # ==== E. JOIN / END STUDS (GEOMETRY-FIRST) ====
        phys_start_d, phys_end_d = self._physical_end_distances(
            wall, p0, direction, length, stud_bottom + STUD_THICKNESS)
        if phys_end_d - phys_start_d < MIN_LENGTH:
            phys_start_d, phys_end_d = (0.0, length)

        run_length = max(0.0, phys_end_d - phys_start_d)
        edge_backset = STUD_THICKNESS * 0.5 + self._wrapped_layer_backset(host_info)
        max_backset = max(0.0, (run_length - STUD_THICKNESS) * 0.5)
        edge_backset = min(edge_backset, max_backset)

        def _line_to_physical(line_dist):
            if length <= 1e-9:
                return (phys_start_d + phys_end_d) * 0.5
            ratio = max(0.0, min(1.0, line_dist / length))
            return phys_start_d + (phys_end_d - phys_start_d) * ratio

        def _has_revit_join_at_end(end_index):
            try:
                return bool(DB.WallUtils.IsWallJoinAllowedAtEnd(wall, end_index))
            except Exception:
                return True

        def _place_join_stud_piece(pos):
            """Place one join/end stud split by mid-plate tiers."""
            pos = max(phys_start_d, min(phys_end_d, pos))
            if self._near(pos, occupied, STUD_THICKNESS * 0.9):
                return True
            if self._in_opening(pos, openings):
                return False

            local_top = self._top_of_stud_at(top_profile, pos)
            tier_bounds = [stud_bottom] + list(mid_plate_zs) + [local_top]
            last_tier = len(tier_bounds) - 2
            placed_any = False
            for ti in range(len(tier_bounds) - 1):
                t_bot = tier_bounds[ti]
                t_top = tier_bounds[ti + 1]
                if ti > 0:
                    t_bot += PLATE_THICKNESS
                t_h = t_top - t_bot
                if t_h >= MIN_LENGTH:
                    sl = _stud_columns if ti == last_tier else None
                    self._place_stud(p0, direction, pos, t_bot,
                                     t_h, stud_sym, level, wall_angle, base_z,
                                     sl)
                    placed_any = True
            if placed_any:
                occupied.add(round(pos, 4))
            return placed_any

        def _place_end_stud_recipe(end_index, target_pos):
            """Deterministic inward search from physical boundary."""
            step = STUD_THICKNESS if end_index == 0 else -STUD_THICKNESS
            for i in range(0, 7):
                cand = target_pos + step * i
                cand = max(phys_start_d, min(phys_end_d, cand))
                if _place_join_stud_piece(cand):
                    return

        def _place_intersection_stud(target_pos):
            if self._in_opening(target_pos, openings):
                return
            for offset in (0.0, STUD_THICKNESS, -STUD_THICKNESS):
                if _place_join_stud_piece(target_pos + offset):
                    return

        for end_index in (0, 1):
            has_revit_join = _has_revit_join_at_end(end_index)

            if not has_revit_join:
                free_target = (
                    phys_start_d + edge_backset
                    if end_index == 0
                    else phys_end_d - edge_backset
                )
                _place_end_stud_recipe(end_index, free_target)
                continue

            end_plan = None
            if join_plan is not None:
                try:
                    end_plan = join_plan.ends.get(end_index)
                except Exception:
                    end_plan = None

            line_positions = []
            if end_plan is not None and end_plan.has_join:
                if end_plan.positions:
                    line_positions = list(end_plan.positions)

            if not line_positions:
                free_target = (
                    phys_start_d + edge_backset
                    if end_index == 0
                    else phys_end_d - edge_backset
                )
                _place_end_stud_recipe(end_index, free_target)
                continue

            for line_pos in line_positions:
                if line_pos < -1e-9 or line_pos > length + 1e-9:
                    continue
                _place_end_stud_recipe(end_index, _line_to_physical(line_pos))

        if join_plan is not None:
            for intersection in getattr(join_plan, "intersections", []) or []:
                raw_positions = list(getattr(intersection, "positions", []) or [])
                if not raw_positions:
                    distance = getattr(intersection, "distance", None)
                    if distance is None:
                        continue
                    raw_positions = [
                        distance - (STUD_THICKNESS * 0.5),
                        distance + (STUD_THICKNESS * 0.5),
                    ]

                for line_pos in raw_positions:
                    if line_pos <= STUD_THICKNESS:
                        continue
                    if line_pos >= length - STUD_THICKNESS:
                        continue
                    _place_intersection_stud(_line_to_physical(line_pos))

        # ==== F. INFILL STUDS (tier-based at mid plates) ====
        spacing = self.config.stud_spacing_ft
        if spacing > 0:
            d = spacing
            while d < length - STUD_THICKNESS * 0.5:
                rd = round(d, 4)
                if not self._in_opening(d, openings) and not self._near(d, occupied):
                    local_stud_top = self._top_of_stud_at(top_profile, d)
                    tier_bounds = [stud_bottom] + list(mid_plate_zs) + [local_stud_top]
                    last_tier = len(tier_bounds) - 2
                    for ti in range(len(tier_bounds) - 1):
                        t_bot = tier_bounds[ti]
                        t_top = tier_bounds[ti + 1]
                        if ti > 0:
                            t_bot += PLATE_THICKNESS
                        t_h = t_top - t_bot
                        if t_h >= MIN_LENGTH:
                            sl = _stud_columns if ti == last_tier else None
                            self._place_stud(p0, direction, d, t_bot,
                                             t_h, stud_sym, level, wall_angle, base_z,
                                             sl)
                    occupied.add(rd)
                d += spacing

        # ==== G. MID-HEIGHT BLOCKING ====
        sorted_occ = sorted(occupied)
        for i in range(len(sorted_occ) - 1):
            d_left  = sorted_occ[i]
            d_right = sorted_occ[i + 1]
            gap = d_right - d_left
            if gap < MIN_LENGTH or gap > spacing * 1.5:
                continue
            mid_d = (d_left + d_right) / 2.0
            local_stud_top = self._top_of_stud_at(top_profile, mid_d)
            stud_h = local_stud_top - stud_bottom
            if stud_h > BLOCKING_MAX_HEIGHT and plate_sym:
                block_z = stud_bottom + stud_h / 2.0
                bs = p0 + direction * (d_left + STUD_THICKNESS * 0.5) + DB.XYZ(0, 0, block_z - base_z)
                be = p0 + direction * (d_right - STUD_THICKNESS * 0.5) + DB.XYZ(0, 0, block_z - base_z)
                self._place(place_beam(self.doc, bs, be, plate_sym, level, PLATE_ROT))

        # ==== H. ATTACH STUD TOPS TO PLATES ====
        self._attach_studs_to_plates(_stud_columns, _plate_beams)

        return True

    # ------------------------------------------------------------------
    def _place(self, inst):
        if inst:
            self.placed.append(inst)

    def _place_stud(self, p0, direction, d, bottom_z, height,
                    stud_sym, level, wall_angle, base_z,
                    stud_list=None):
        pt = p0 + direction * d + DB.XYZ(0, 0, bottom_z - base_z)
        inst = place_column(self.doc, pt, height, stud_sym, level, wall_angle)
        if inst:
            self.placed.append(inst)
            if stud_list is not None:
                stud_list.append((inst, d))

    def _place_cripples_oc(self, p0, direction, left, right, bottom_z,
                           height, oc_spacing, stud_sym, level, wall_angle,
                           base_z, occupied, stud_list=None):
        """Place cripple studs aligned to the wall's global OC grid."""
        if oc_spacing <= 0 or height < MIN_LENGTH:
            return
        # Walk the global OC grid and place cripples where they fall in the opening
        d = oc_spacing
        while d < right + STUD_THICKNESS:
            if left - STUD_THICKNESS < d < right + STUD_THICKNESS:
                if not self._near(d, occupied):
                    self._place_stud(p0, direction, d, bottom_z, height,
                                     stud_sym, level, wall_angle, base_z,
                                     stud_list)
            if d > right + STUD_THICKNESS:
                break
            d += oc_spacing

    def _attach_studs_to_plates(self, stud_columns, plate_beams):
        """Attach each stud's top to the angled plate beam above it.

        Matches by position along wall (d) — if stud d falls within
        the plate segment's d-range, attach it.
        """
        if not stud_columns or not plate_beams:
            return
        for stud_inst, stud_d in stud_columns:
            best_plate = None
            for plate_inst, p_d0, p_d1, p_is_angled in plate_beams:
                if not p_is_angled:
                    continue
                if p_d0 - STUD_THICKNESS <= stud_d <= p_d1 + STUD_THICKNESS:
                    best_plate = plate_inst
                    break
            if best_plate is None:
                continue

            # Skip invalid or already-joined pairs to avoid Revit join warnings/errors.
            try:
                if not DB.ColumnAttachment.IsValidColumn(stud_inst):
                    continue
            except Exception:
                continue
            try:
                if not DB.ColumnAttachment.IsValidTarget(stud_inst, best_plate):
                    continue
            except Exception:
                continue
            try:
                if DB.JoinGeometryUtils.AreElementsJoined(self.doc, stud_inst, best_plate):
                    continue
            except Exception:
                pass
            try:
                existing = DB.ColumnAttachment.GetColumnAttachment(
                    stud_inst, best_plate.Id)
                if existing is not None and existing.IsValidObject:
                    continue
            except Exception:
                pass

            try:
                DB.ColumnAttachment.AddColumnAttachment(
                    self.doc,
                    stud_inst,
                    best_plate,
                    1,  # 1 = top
                    getattr(DB.ColumnAttachmentCutStyle, "None"),
                    DB.ColumnAttachmentJustification.Minimum,
                    0.0,
                )
            except Exception:
                pass

    def _stud_height_at(self, top_profile, d, stud_bottom):
        top_z = self._interpolate_z(top_profile, d)
        return top_z - PLATE_THICKNESS * self.config.top_plate_count - stud_bottom

    def _top_of_stud_at(self, top_profile, d):
        top_z = self._interpolate_z(top_profile, d)
        return top_z - PLATE_THICKNESS * self.config.top_plate_count

    # ------------------------------------------------------------------
    #  Detect where other walls intersect this wall (T / + junctions)
    # ------------------------------------------------------------------
    def _find_wall_intersections(self, wall, p0, direction, length):
        """Return list of distances along wall where other walls join.

        Uses the wall's joined elements at each end plus any walls
        whose endpoint lies on this wall's location line.
        """
        doc = self.doc
        wall_id = wall.Id
        dists = []

        # Collect all walls on the same level
        try:
            walls = (
                DB.FilteredElementCollector(doc)
                .OfClass(DB.Wall)
                .WhereElementIsNotElementType()
            )
        except Exception:
            return dists

        for other in walls:
            try:
                if other.Id == wall_id:
                    continue
                other_loc = other.Location
                if other_loc is None:
                    continue
                other_curve = other_loc.Curve
                if not isinstance(other_curve, DB.Line):
                    continue

                # Check if either endpoint of the other wall lies on this wall
                for ei in (0, 1):
                    ep = other_curve.GetEndPoint(ei)
                    # Project onto this wall's direction
                    vec = ep - p0
                    d_along = vec.DotProduct(direction)
                    # Perpendicular distance to this wall's line
                    parallel_pt = p0 + direction * d_along
                    perp_dist = ep.DistanceTo(parallel_pt)

                    # Must be close to this wall (within wall thickness)
                    # and not at the wall ends (those are corners, not T-junctions)
                    if (perp_dist < 1.0 and
                        d_along > STUD_THICKNESS * 3 and
                        d_along < length - STUD_THICKNESS * 3):
                        # Avoid duplicates
                        is_dup = False
                        for existing in dists:
                            if abs(existing - d_along) < STUD_THICKNESS * 2:
                                is_dup = True
                                break
                        if not is_dup:
                            dists.append(d_along)
            except Exception:
                continue
        return dists

    def _connected_wall_ends(self, wall, p0, direction, length):
        """Return (start_connected, end_connected) booleans.

        True when another wall's endpoint meets this wall's start/end,
        forming an L or T corner.  Used to decide whether to add a
        second corner stud (only needed at free ends).
        """
        doc = self.doc
        wall_id = wall.Id
        ep_start = p0
        ep_end = p0 + direction * length
        start_conn = False
        end_conn = False
        tol = STUD_THICKNESS * 2

        try:
            walls = (
                DB.FilteredElementCollector(doc)
                .OfClass(DB.Wall)
                .WhereElementIsNotElementType()
            )
        except Exception:
            return (False, False)

        for other in walls:
            try:
                if other.Id == wall_id:
                    continue
                oloc = other.Location
                if oloc is None:
                    continue
                oc = oloc.Curve
                if not isinstance(oc, DB.Line):
                    continue
                for ei in (0, 1):
                    opt = oc.GetEndPoint(ei)
                    if opt.DistanceTo(ep_start) < tol:
                        start_conn = True
                    if opt.DistanceTo(ep_end) < tol:
                        end_conn = True
                if start_conn and end_conn:
                    break
            except Exception:
                continue
        return (start_conn, end_conn)

    #  Detect ALL openings in a wall: doors, windows, curtain walls,
    #  rectangular openings (voids), and generic inserts.
    # ------------------------------------------------------------------
    def _get_openings(self, wall, p0, direction, base_z):
        doc = self.doc
        wall_id = wall.Id
        openings = []

        # --- 1. Doors and Windows via category collectors ---
        for bic, is_window in (
            (DB.BuiltInCategory.OST_Doors, False),
            (DB.BuiltInCategory.OST_Windows, True),
        ):
            try:
                collector = (
                    DB.FilteredElementCollector(doc)
                    .OfCategory(bic)
                    .WhereElementIsNotElementType()
                )
            except Exception:
                continue

            for elem in collector:
                try:
                    if not isinstance(elem, DB.FamilyInstance):
                        continue
                    host = elem.Host
                    if host is None or host.Id != wall_id:
                        continue

                    loc = elem.Location
                    if loc is None:
                        continue
                    try:
                        c_pt = loc.Point
                    except Exception:
                        continue

                    c_dist = (c_pt - p0).DotProduct(direction)
                    ow = self._read_opening_dim(elem, "width")
                    oh = self._read_opening_dim(elem, "height")

                    osill = 0.0
                    if is_window:
                        osill = self._read_sill_height(elem)

                    openings.append({
                        "left":  c_dist - ow / 2.0,
                        "right": c_dist + ow / 2.0,
                        "head_z": base_z + osill + oh,
                        "sill_z": base_z + osill,
                        "is_window": is_window,
                    })
                except Exception:
                    continue

        # --- 2. Rectangular openings (voids) and embedded curtain walls ---
        # Use FindInserts but ONLY accept Opening elements and curtain walls.
        try:
            insert_ids = wall.FindInserts(True, False, True, True)
            if insert_ids:
                wall_h_param = wall.get_Parameter(
                    DB.BuiltInParameter.WALL_USER_HEIGHT_PARAM)
                wall_height = wall_h_param.AsDouble() if (
                    wall_h_param and wall_h_param.HasValue) else 8.0

                for eid in insert_ids:
                    try:
                        elem = doc.GetElement(eid)
                        if elem is None:
                            continue

                        # Rectangular Opening (void cut)
                        if isinstance(elem, DB.Opening):
                            bp = elem.BoundaryRect
                            if bp and len(bp) >= 2:
                                mn = bp[0]
                                mx = bp[1]
                                d_min = (mn - p0).DotProduct(direction)
                                d_max = (mx - p0).DotProduct(direction)
                                left = min(d_min, d_max)
                                right = max(d_min, d_max)
                                z_lo = min(mn.Z, mx.Z)
                                z_hi = max(mn.Z, mx.Z)
                                # Skip if already captured by door/window
                                is_dup = False
                                for ex in openings:
                                    if (abs(ex["left"] - left) < STUD_THICKNESS and
                                            abs(ex["right"] - right) < STUD_THICKNESS):
                                        is_dup = True
                                        break
                                if not is_dup:
                                    openings.append({
                                        "left": left,
                                        "right": right,
                                        "head_z": z_hi,
                                        "sill_z": z_lo,
                                        "is_window": z_lo > base_z + PLATE_THICKNESS,
                                    })
                            continue

                        # Embedded curtain wall
                        if isinstance(elem, DB.Wall):
                            try:
                                wk = elem.WallType.Kind
                                if wk != DB.WallKind.Curtain:
                                    continue
                            except Exception:
                                continue
                            loc = elem.Location
                            if loc is None:
                                continue
                            try:
                                ec = loc.Curve
                                ep0 = ec.GetEndPoint(0)
                                ep1 = ec.GetEndPoint(1)
                            except Exception:
                                continue
                            d0 = (ep0 - p0).DotProduct(direction)
                            d1 = (ep1 - p0).DotProduct(direction)
                            left = min(d0, d1)
                            right = max(d0, d1)
                            if right - left < MIN_LENGTH:
                                continue
                            cw_base = base_z
                            try:
                                bo = elem.get_Parameter(
                                    DB.BuiltInParameter.WALL_BASE_OFFSET)
                                if bo and bo.HasValue:
                                    cw_base += bo.AsDouble()
                            except Exception:
                                pass
                            cw_h_param = elem.get_Parameter(
                                DB.BuiltInParameter.WALL_USER_HEIGHT_PARAM)
                            cw_h = cw_h_param.AsDouble() if (
                                cw_h_param and cw_h_param.HasValue) else wall_height
                            is_dup = False
                            for ex in openings:
                                if (abs(ex["left"] - left) < STUD_THICKNESS and
                                        abs(ex["right"] - right) < STUD_THICKNESS):
                                    is_dup = True
                                    break
                            if not is_dup:
                                openings.append({
                                    "left": left,
                                    "right": right,
                                    "head_z": cw_base + cw_h,
                                    "sill_z": cw_base,
                                    "is_window": cw_base > base_z + PLATE_THICKNESS,
                                })
                    except Exception:
                        continue
        except Exception:
            pass

        return openings

    @staticmethod
    def _opening_dicts_from_host_info(host_info):
        openings = []
        base_z = host_info.base_elevation
        for opening in getattr(host_info, "openings", []):
            openings.append({
                "left": opening.left_edge,
                "right": opening.right_edge,
                "head_z": base_z + opening.head_height,
                "sill_z": base_z + opening.sill_height,
                "is_window": opening.is_window,
            })
        return openings

    @staticmethod
    def _read_opening_dim(elem, dim_kind):
        """Read width or height from a door/window, trying multiple sources."""
        if dim_kind == "width":
            bips = tuple(
                getattr(DB.BuiltInParameter, name, None)
                for name in (
                    "FAMILY_ROUGH_WIDTH_PARAM",
                    "DOOR_ROUGH_WIDTH",
                    "WINDOW_ROUGH_WIDTH",
                    "FAMILY_WIDTH_PARAM",
                    "GENERIC_WIDTH",
                )
                if getattr(DB.BuiltInParameter, name, None) is not None
            )
            names = ("Rough Width", "Rough Opening Width", "Width")
            default = 3.0
        else:
            bips = tuple(
                getattr(DB.BuiltInParameter, name, None)
                for name in (
                    "FAMILY_ROUGH_HEIGHT_PARAM",
                    "DOOR_ROUGH_HEIGHT",
                    "WINDOW_ROUGH_HEIGHT",
                    "FAMILY_HEIGHT_PARAM",
                    "GENERIC_HEIGHT",
                )
                if getattr(DB.BuiltInParameter, name, None) is not None
            )
            names = ("Rough Height", "Rough Opening Height", "Height")
            default = 6.67

        # Instance built-in parameters
        for bip in bips:
            try:
                p = elem.get_Parameter(bip)
                if p and p.HasValue and p.AsDouble() > 0.1:
                    return p.AsDouble()
            except Exception:
                pass
        # Instance named parameters
        for n in names:
            try:
                p = elem.LookupParameter(n)
                if p and p.HasValue and p.AsDouble() > 0.1:
                    return p.AsDouble()
            except Exception:
                pass
        # Type parameters
        sym = elem.Symbol if hasattr(elem, "Symbol") else None
        if sym:
            for bip in bips:
                try:
                    p = sym.get_Parameter(bip)
                    if p and p.HasValue and p.AsDouble() > 0.1:
                        return p.AsDouble()
                except Exception:
                    pass
            for n in names:
                try:
                    p = sym.LookupParameter(n)
                    if p and p.HasValue and p.AsDouble() > 0.1:
                        return p.AsDouble()
                except Exception:
                    pass
        return default

    @staticmethod
    def _read_sill_height(elem):
        """Read window sill height from floor level."""
        for bip in (DB.BuiltInParameter.INSTANCE_SILL_HEIGHT_PARAM,):
            try:
                p = elem.get_Parameter(bip)
                if p and p.HasValue:
                    return p.AsDouble()
            except Exception:
                pass
        for n in ("Sill Height",):
            try:
                p = elem.LookupParameter(n)
                if p and p.HasValue:
                    return p.AsDouble()
            except Exception:
                pass
        return 3.0  # 3 ft default sill height

    @staticmethod
    def _split_segments(start, end, gaps):
        segments = [(start, end)]
        for gs, ge in gaps:
            new_segs = []
            for ss, se in segments:
                if ge <= ss or gs >= se:
                    new_segs.append((ss, se))
                    continue
                if gs > ss:
                    new_segs.append((ss, gs))
                if ge < se:
                    new_segs.append((ge, se))
            segments = new_segs
        return [(s, e) for s, e in segments if e - s >= MIN_LENGTH]

    @staticmethod
    def _in_opening(dist, openings):
        for op in openings:
            # Exclude zone covers king+jack studs (up to 1.5 stud widths outside RO)
            if (op["left"] - STUD_THICKNESS * 3) <= dist <= (op["right"] + STUD_THICKNESS * 3):
                return True
        return False

    @staticmethod
    def _near(dist, occupied, tol=None):
        if tol is None:
            tol = STUD_THICKNESS
        rd = round(dist, 4)
        for o in occupied:
            if abs(rd - o) < tol:
                return True
        return False

    @staticmethod
    def _wrapped_layer_backset(host_info):
        if host_info is None:
            return 0.0
        base_info = getattr(host_info, "wall_info", None)
        wall_width = getattr(base_info, "width", 0.0)
        target_layer = getattr(host_info, "target_layer", None)
        target_width = getattr(target_layer, "width", 0.0)
        if wall_width <= 1e-9 or target_width <= 1e-9:
            return 0.0
        return max(0.0, (wall_width - target_width) * 0.5)

    # ------------------------------------------------------------------
    #  Wall top profile — read the wall's own shape
    # ------------------------------------------------------------------
    def _wall_top_profile(self, wall, p0, direction, base_z, length):
        """Return [(d, z), ...] for the wall's top edge.

        1. Edited-profile walls have a Sketch — read its curves.
        2. Shoot vertical rays through the wall solid to detect the
           actual shape (gable / attached-to-roof / any non-rectangular).
        3. If rays fail, BoundingBox gives the correct flat top Z.
        """
        # BoundingBox top — always correct for the peak Z
        bb_top = base_z + 8.0
        try:
            bb = wall.get_BoundingBox(None)
            if bb is not None and bb.Max.Z > base_z + 0.01:
                bb_top = bb.Max.Z
        except Exception:
            pass

        # 1. Try sketch profile (manually edited profile walls)
        profile = self._profile_from_sketch(wall, p0, direction)
        if profile and len(profile) >= 2:
            return profile

        # 2. Ray-based sampling — detects gable/sloped/attached shapes
        profile = self._profile_from_rays(wall, p0, direction,
                                          base_z, length, bb_top)
        if profile and len(profile) >= 2:
            return profile

        # 3. Flat fallback from BoundingBox
        return [(0.0, bb_top), (length, bb_top)]

    def _profile_from_rays(self, wall, p0, direction, base_z, length, bb_top):
        """Shoot vertical rays along wall length, return top Z at each."""
        try:
            opts = DB.Options()
            opts.ComputeReferences = False
            opts.DetailLevel = DB.ViewDetailLevel.Fine
            geom_elem = wall.get_Geometry(opts)
        except Exception:
            return None

        solids = []
        for gobj in geom_elem:
            if isinstance(gobj, DB.Solid) and gobj.Volume > 0:
                solids.append(gobj)
            elif isinstance(gobj, DB.GeometryInstance):
                try:
                    for sub in gobj.GetInstanceGeometry():
                        if isinstance(sub, DB.Solid) and sub.Volume > 0:
                            solids.append(sub)
                except Exception:
                    pass
        if not solids:
            return None

        ray_top = bb_top + 5.0
        n = max(10, int(length / 0.5))
        profile = []
        for i in range(n + 1):
            d = length * i / n
            pt = p0 + direction * d
            try:
                ray = DB.Line.CreateBound(
                    DB.XYZ(pt.X, pt.Y, base_z - 1.0),
                    DB.XYZ(pt.X, pt.Y, ray_top),
                )
            except Exception:
                continue
            max_z = base_z
            for solid in solids:
                try:
                    result = solid.IntersectWithCurve(
                        ray, DB.SolidCurveIntersectionOptions())
                    for si in range(result.SegmentCount):
                        ep = result.GetCurveSegment(si).GetEndPoint(1)
                        if ep.Z > max_z:
                            max_z = ep.Z
                except Exception:
                    continue
            if max_z > base_z + 0.5:
                profile.append((d, max_z))
        return profile if len(profile) >= 2 else None

    def _profile_from_sketch(self, wall, p0, direction):
        """Extract the top profile from the wall's profile Sketch."""
        try:
            sketch_id = wall.SketchId
            raw_id = getattr(sketch_id, "IntegerValue",
                             getattr(sketch_id, "Value", -1))
            if raw_id == -1:
                return None
            sketch = self.doc.GetElement(sketch_id)
            if sketch is None:
                return None
        except Exception:
            return None

        # Collect all curve points from the sketch profile
        points = []
        try:
            for curve_arr in sketch.Profile:
                for curve in curve_arr:
                    for pt in curve.Tessellate():
                        d = (pt - p0).DotProduct(direction)
                        points.append((d, pt.Z))
        except Exception:
            return None

        if len(points) < 3:
            return None

        # Separate top half of points (above mid-height)
        min_z = min(z for _, z in points)
        max_z = max(z for _, z in points)
        mid_z = (min_z + max_z) / 2.0

        # Bucket by distance, keep highest Z at each position
        bucket_sz = 1.0 / 12.0
        buckets = {}
        for d, z in points:
            if z < mid_z:
                continue
            k = round(d / bucket_sz)
            if k not in buckets or z > buckets[k][1]:
                buckets[k] = (d, z)

        profile = sorted(buckets.values())
        return profile if len(profile) >= 2 else None

    @staticmethod
    def _simplify_profile(profile, tol=1.0 / 12.0):
        if len(profile) <= 2:
            return list(profile)
        result = [profile[0]]
        for i in range(1, len(profile) - 1):
            dp, zp = result[-1]
            dc, zc = profile[i]
            dn, zn = profile[i + 1]
            span = dn - dp
            if abs(span) < 1e-9:
                continue
            t = (dc - dp) / span
            if abs(zc - (zp + t * (zn - zp))) > tol:
                result.append(profile[i])
        result.append(profile[-1])
        return result

    @staticmethod
    def _interpolate_z(profile, d):
        if not profile:
            return 0.0
        if d <= profile[0][0]:
            return profile[0][1]
        if d >= profile[-1][0]:
            return profile[-1][1]
        for i in range(len(profile) - 1):
            d0, z0 = profile[i]
            d1, z1 = profile[i + 1]
            if d0 <= d <= d1:
                t = (d - d0) / (d1 - d0) if abs(d1 - d0) > 1e-9 else 0.0
                return z0 + t * (z1 - z0)
        return profile[-1][1]

    def _physical_end_distances(self, wall, p0, direction, length, sample_z):
        """Get physical wall run limits from 3D solid intersections.

        Casts a horizontal ray along the framing line at stud-zone height.
        This avoids relying on wall location-curve endpoints when walls join.
        """
        try:
            opts = DB.Options()
            opts.ComputeReferences = False
            opts.DetailLevel = DB.ViewDetailLevel.Fine
            geom_elem = wall.get_Geometry(opts)
        except Exception:
            return (0.0, length)

        solids = []
        for gobj in geom_elem:
            if isinstance(gobj, DB.Solid) and gobj.Volume > 0:
                solids.append(gobj)
            elif isinstance(gobj, DB.GeometryInstance):
                try:
                    for sub in gobj.GetInstanceGeometry():
                        if isinstance(sub, DB.Solid) and sub.Volume > 0:
                            solids.append(sub)
                except Exception:
                    pass
        if not solids:
            return (0.0, length)

        z_shift = sample_z - p0.Z
        ray_origin = p0 + DB.XYZ(0, 0, z_shift)
        ray_pad = max(2.0, length * 0.25, STUD_THICKNESS * 8.0)
        ray_start = ray_origin - direction * ray_pad
        ray_end = ray_origin + direction * (length + ray_pad)

        try:
            ray = DB.Line.CreateBound(ray_start, ray_end)
        except Exception:
            return (0.0, length)

        hit_min = None
        hit_max = None
        for solid in solids:
            try:
                result = solid.IntersectWithCurve(
                    ray, DB.SolidCurveIntersectionOptions())
            except Exception:
                continue
            try:
                seg_count = result.SegmentCount
            except Exception:
                seg_count = 0
            for si in range(seg_count):
                try:
                    seg = result.GetCurveSegment(si)
                    p_a = seg.GetEndPoint(0)
                    p_b = seg.GetEndPoint(1)
                except Exception:
                    continue
                d_a = (p_a - p0).DotProduct(direction)
                d_b = (p_b - p0).DotProduct(direction)
                lo = min(d_a, d_b)
                hi = max(d_a, d_b)
                if hit_min is None or lo < hit_min:
                    hit_min = lo
                if hit_max is None or hi > hit_max:
                    hit_max = hi

        if hit_min is None or hit_max is None:
            return (0.0, length)

        min_d = max(0.0, min(length, hit_min))
        max_d = max(0.0, min(length, hit_max))
        if max_d - min_d < MIN_LENGTH:
            return (0.0, length)
        return (min_d, max_d)


# =========================================================================
#  Config persistence
# =========================================================================

def _save_last_config(cfg):
    try:
        d = os.path.dirname(_CFG_PATH)
        if not os.path.exists(d):
            os.makedirs(d)
        cfg.save(_CFG_PATH)
    except Exception:
        pass


def _load_last_config():
    """Load last-used config; fall back to global defaults from WF Settings."""
    try:
        if os.path.exists(_CFG_PATH):
            return FramingConfig.load(_CFG_PATH)
    except Exception:
        pass
    # Fall back to global defaults set via WF Settings
    try:
        global_path = os.path.join(_lib_dir, "wf_defaults.json")
        if os.path.exists(global_path):
            return FramingConfig.load(global_path)
    except Exception:
        pass
    return None


# =========================================================================
#  WPF Dialog — remembers last state
# =========================================================================

class FrameWallDialog(forms.WPFWindow):
    def __init__(self, xaml_file, column_types, framing_types, last_cfg=None):
        forms.WPFWindow.__init__(self, xaml_file)
        self._column_types = column_types
        self._framing_types = framing_types
        self._last_cfg = last_cfg
        self.result_config = None
        self._populate_combos()
        self._restore_state()
        self.rb_custom.Checked += self._on_custom_checked
        self.rb_custom.Unchecked += self._on_custom_unchecked
        self.rb_16oc.Checked += self._on_custom_unchecked
        self.rb_24oc.Checked += self._on_custom_unchecked
        self.chk_mid_plates.Checked += self._on_mid_plates_checked
        self.chk_mid_plates.Unchecked += self._on_mid_plates_unchecked

    def _populate_combos(self):
        self.cb_stud_type.Items.Clear()
        for t in self._column_types:
            self.cb_stud_type.Items.Add(t)
        if self.cb_stud_type.Items.Count > 0:
            self.cb_stud_type.SelectedIndex = 0

        for combo in (self.cb_plate_type, self.cb_header_type):
            combo.Items.Clear()
            for t in self._framing_types:
                combo.Items.Add(t)
            if combo.Items.Count > 0:
                combo.SelectedIndex = 0

        self.cb_wall_layer_mode.Items.Clear()
        self.cb_wall_layer_mode.Items.Add("Core center")
        self.cb_wall_layer_mode.Items.Add("Structural layer")
        self.cb_wall_layer_mode.Items.Add("Thickest layer")
        self.cb_wall_layer_mode.SelectedIndex = 1

    def _restore_state(self):
        cfg = self._last_cfg
        if cfg is None:
            return
        if cfg.stud_family_name and cfg.stud_type_name:
            self._select_combo(self.cb_stud_type,
                               "{0} : {1}".format(cfg.stud_family_name, cfg.stud_type_name))
        pf = cfg.bottom_plate_family_name or cfg.stud_family_name
        pt = cfg.bottom_plate_type_name or cfg.stud_type_name
        if pf and pt:
            self._select_combo(self.cb_plate_type, "{0} : {1}".format(pf, pt))
        hf = cfg.header_family_name or cfg.stud_family_name
        ht = cfg.header_type_name or cfg.stud_type_name
        if hf and ht:
            self._select_combo(self.cb_header_type, "{0} : {1}".format(hf, ht))
        if cfg.stud_spacing == SPACING_16OC:
            self.rb_16oc.IsChecked = True
        elif cfg.stud_spacing == SPACING_24OC:
            self.rb_24oc.IsChecked = True
        else:
            self.rb_custom.IsChecked = True
            self.tb_custom_spacing.Text = str(cfg.stud_spacing)
            self.tb_custom_spacing.IsEnabled = True
        if cfg.top_plate_count == 1:
            self.rb_single_top.IsChecked = True
        else:
            self.rb_double_top.IsChecked = True
        self.chk_king_studs.IsChecked = cfg.include_king_studs
        self.chk_jack_studs.IsChecked = cfg.include_jack_studs
        self.chk_cripple_studs.IsChecked = cfg.include_cripple_studs
        self.chk_mid_plates.IsChecked = bool(getattr(cfg, "include_mid_plates", True))
        self.tb_mid_plate_height.Text = str(
            float(getattr(cfg, "mid_plate_interval_ft", MID_PLATE_INTERVAL))
        )
        self.tb_mid_plate_height.IsEnabled = bool(self.chk_mid_plates.IsChecked)
        self.chk_support_top.IsChecked = (
            getattr(cfg, "wall_base_mode", WALL_BASE_MODE_WALL) == WALL_BASE_MODE_SUPPORT_TOP
        )
        mode_map = {"core_center": 0, "structural": 1, "thickest": 2}
        self.cb_wall_layer_mode.SelectedIndex = mode_map.get(cfg.wall_layer_mode, 1)

    @staticmethod
    def _select_combo(combo, target):
        for i in range(combo.Items.Count):
            if str(combo.Items[i]) == target:
                combo.SelectedIndex = i
                return

    def _on_custom_checked(self, sender, args):
        self.tb_custom_spacing.IsEnabled = True

    def _on_custom_unchecked(self, sender, args):
        self.tb_custom_spacing.IsEnabled = False

    def _on_mid_plates_checked(self, sender, args):
        self.tb_mid_plate_height.IsEnabled = True

    def _on_mid_plates_unchecked(self, sender, args):
        self.tb_mid_plate_height.IsEnabled = False

    def btn_ok_click(self, sender, args):
        cfg = FramingConfig()
        if self.rb_16oc.IsChecked:
            cfg.stud_spacing = SPACING_16OC
        elif self.rb_24oc.IsChecked:
            cfg.stud_spacing = SPACING_24OC
        else:
            try:
                cfg.stud_spacing = float(self.tb_custom_spacing.Text)
            except Exception:
                forms.alert("Invalid custom spacing value.")
                return

        stud_sel = self.cb_stud_type.SelectedItem
        if not stud_sel:
            forms.alert("Select a stud (column) family.", title="Error")
            return
        fam, typ = parse_family_type_label(str(stud_sel))
        cfg.stud_family_name, cfg.stud_type_name = fam, typ

        if self.cb_plate_type.SelectedItem:
            f, t = parse_family_type_label(str(self.cb_plate_type.SelectedItem))
            cfg.bottom_plate_family_name, cfg.bottom_plate_type_name = f, t
            cfg.top_plate_family_name, cfg.top_plate_type_name = f, t

        if self.cb_header_type.SelectedItem:
            f, t = parse_family_type_label(str(self.cb_header_type.SelectedItem))
            cfg.header_family_name, cfg.header_type_name = f, t

        cfg.top_plate_count = 1 if self.rb_single_top.IsChecked else 2
        cfg.include_king_studs = bool(self.chk_king_studs.IsChecked)
        cfg.include_jack_studs = bool(self.chk_jack_studs.IsChecked)
        cfg.include_cripple_studs = bool(self.chk_cripple_studs.IsChecked)
        cfg.include_mid_plates = bool(self.chk_mid_plates.IsChecked)
        try:
            cfg.mid_plate_interval_ft = float(self.tb_mid_plate_height.Text)
        except Exception:
            forms.alert("Invalid mid-plate interval value.")
            return
        if cfg.mid_plate_interval_ft <= 0.0:
            forms.alert("Mid-plate interval must be greater than 0.")
            return

        if bool(self.chk_support_top.IsChecked):
            cfg.wall_base_mode = WALL_BASE_MODE_SUPPORT_TOP
        else:
            cfg.wall_base_mode = WALL_BASE_MODE_WALL
        cfg.wall_base_override_z = None
        cfg.wall_base_support_element_id = None

        mode_idx = self.cb_wall_layer_mode.SelectedIndex
        cfg.wall_layer_mode = [LAYER_MODE_CORE_CENTER,
                               LAYER_MODE_STRUCTURAL,
                               LAYER_MODE_THICKEST][mode_idx]

        self.result_config = cfg
        self.DialogResult = True
        self.Close()

    def btn_cancel_click(self, sender, args):
        self.DialogResult = False
        self.Close()


# =========================================================================
#  Entry point
# =========================================================================

class _WallFilter(ISelectionFilter):
    def AllowElement(self, el):
        return isinstance(el, DB.Wall)
    def AllowReference(self, ref, pt):
        return False


class _WallSupportFilter(ISelectionFilter):
    """Allow floor-like hosts that can define a wall framing baseline."""

    def AllowElement(self, el):
        if isinstance(el, DB.Floor):
            return True
        if isinstance(el, DB.RoofBase):
            return True
        category = getattr(el, "Category", None)
        if category is None:
            return False
        try:
            return int(category.Id.IntegerValue) == int(DB.BuiltInCategory.OST_StructuralFoundation)
        except Exception:
            return False

    def AllowReference(self, ref, pt):
        return False


def _pick_wall_support_element(doc):
    """Prompt user to pick the support host used for wall bottom-plate baseline."""
    try:
        picked = revit.uidoc.Selection.PickObject(
            ObjectType.Element,
            _WallSupportFilter(),
            "Pick support host (floor/roof/foundation) for wall base",
        )
    except Exception:
        return None

    if picked is None:
        return None
    try:
        return doc.GetElement(picked.ElementId)
    except Exception:
        return None


def main():
    doc = revit.doc

    framing_types = get_available_types_flat(doc)
    column_types  = get_column_types_flat(doc)

    if not column_types:
        forms.alert(
            "No Structural Column families loaded.\n\n"
            "Studs require a Structural Column family.\n"
            "Load one via Insert > Load Family.",
            title="Wood Framing",
        )
        return
    if not framing_types:
        forms.alert("No Structural Framing families found.", title="Wood Framing")
        return

    selected = revit.get_selection().elements
    walls = [e for e in selected if isinstance(e, DB.Wall)]
    if not walls:
        try:
            refs = revit.uidoc.Selection.PickObjects(
                ObjectType.Element, _WallFilter(), "Select walls to frame"
            )
            walls = [doc.GetElement(r.ElementId) for r in refs]
        except Exception:
            return
    if not walls:
        return

    last_cfg = _load_last_config()

    xaml_path = script.get_bundle_file("FrameWallConfig.xaml")
    dialog = FrameWallDialog(xaml_path, column_types, framing_types, last_cfg)
    dialog.ShowDialog()

    if dialog.result_config is None:
        return
    config = dialog.result_config

    if getattr(config, "wall_base_mode", WALL_BASE_MODE_WALL) == WALL_BASE_MODE_SUPPORT_TOP:
        support = _pick_wall_support_element(doc)
        if support is None:
            forms.alert(
                "Support host selection was cancelled.",
                title="Wood Framing",
            )
            return
        support_id = getattr(
            support.Id,
            "Value",
            getattr(support.Id, "IntegerValue", None),
        )
        config.wall_base_support_element_id = support_id

    # Persist user options but not transient support element references.
    persist_cfg = FramingConfig.from_dict(config.to_dict())
    persist_cfg.wall_base_support_element_id = None
    persist_cfg.wall_base_override_z = None
    _save_last_config(persist_cfg)

    framer = WallFramer(doc, config)
    total = 0
    with revit.Transaction("WF: Frame Walls"):
        for wall in walls:
            if framer.frame_wall(wall):
                total += 1

    output.print_md(
        "## Wood Framing Complete\n"
        "- **Walls framed:** {0}\n"
        "- **Members placed:** {1}\n".format(total, len(framer.placed))
    )


if __name__ == "__main__":
    main()
