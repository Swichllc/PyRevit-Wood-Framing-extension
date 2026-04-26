# -*- coding: utf-8 -*-
"""Frame Wall — wood framing for selected Revit walls.

Studs (vertical)   -> Structural Column families (OST_StructuralColumns)
Plates / Headers    -> Structural Framing families (OST_StructuralFraming)

Construction sequence per wall:
  1. Wall shape: bottom plates, top plates, and side studs from the side face.
  2. Openings: king studs, jack/trimmer studs, headers, sills, and cripples.
  3. Infill: mid plates, OC studs, and blocking inside the remaining spaces.
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
    LAYER_MODE_STRUCTURAL,
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
from wf_schedule_utils import ensure_bom_parameters, apply_bom_metadata
from wf_tracking import get_tracking_data, tag_instance

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
FRAME_WALL_ENGINE = "interior-face-v2"

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


class _WallBomHostInfo(object):
    """Small host shim so legacy wall placement can stamp BOM metadata."""

    def __init__(self, wall):
        self.kind = "wall"
        self.element = wall
        self.element_id = getattr(wall, "Id", None)


class _WallTrackedMember(object):
    """Minimal member descriptor for legacy wall tracking/BOM backfill."""

    def __init__(self, host_info, member_role):
        self.member_type = member_role
        self.host_kind = getattr(host_info, "kind", None)
        self.host_id = getattr(host_info, "element_id", None)
        target_layer = getattr(host_info, "target_layer", None)
        self.layer_index = getattr(target_layer, "index", None)


# =========================================================================
#  Wall framing engine
# =========================================================================

class WallFramer(object):
    def __init__(self, doc, config):
        self.doc = doc
        self.config = config
        self.placed = []
        self._active_host_info = None

    def frame_wall(self, wall):
        loc_curve = wall.Location
        if loc_curve is None:
            return False
        curve = loc_curve.Curve
        if not isinstance(curve, DB.Line):
            return False

        p0 = curve.GetEndPoint(0)
        p1 = curve.GetEndPoint(1)
        direction = (p1 - p0).Normalize()
        normal = safe_wall_normal(wall, direction)
        if normal is None:
            return False
        wall_angle = math.atan2(direction.Y, direction.X)

        host_info = None
        try:
            host_info = analyze_wall_host(self.doc, wall, self.config)
        except Exception:
            pass
        self._active_host_info = host_info or _WallBomHostInfo(wall)
        if host_info is not None:
            direction = host_info.direction
            normal = host_info.normal
            wall_angle = host_info.wall_info.angle
            level = self.doc.GetElement(host_info.level_id)
            if level is None:
                return False
            base_z = host_info.base_elevation
            p0 = host_info.point_at(0.0, 0.0)
            host_openings = self._opening_dicts_from_host_info(host_info)
        else:
            level = self.doc.GetElement(wall.LevelId)
            if level is None:
                return False
            base_z = level.Elevation
            p_off = wall.get_Parameter(DB.BuiltInParameter.WALL_BASE_OFFSET)
            if p_off and p_off.HasValue:
                base_z += p_off.AsDouble()
            p0 = DB.XYZ(p0.X, p0.Y, base_z)
            host_openings = self._get_openings(wall, p0, direction, base_z)

        face_shape = self._wall_face_shape(
            wall,
            p0,
            direction,
            normal,
            base_z,
        )
        if face_shape is None:
            logger.warning(
                "Wall {0}: side-face outline was not readable; skipped wall framing "
                "instead of using location-curve endpoints.".format(wall.Id)
            )
            return False

        face_end_d = face_shape["start_d"] + face_shape["length"]
        trimmed_openings = self._trim_openings_to_span(
            host_openings,
            face_shape["start_d"],
            face_end_d,
        )
        p0 = p0 + direction * face_shape["start_d"]
        length = face_shape["length"]
        top_profile = face_shape["top_profile"]
        openings = self._merge_opening_sets(face_shape["openings"], trimmed_openings)
        if length < MIN_LENGTH:
            return False

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

        # ==== 1A. WALL SHAPE - BOTTOM PLATES ====
        for i in range(self.config.bottom_plate_count):
            pz = base_z + i * PLATE_THICKNESS + PLATE_THICKNESS / 2.0
            for seg_s, seg_e in plate_segs:
                s = p0 + direction * seg_s + DB.XYZ(0, 0, pz - base_z)
                e = p0 + direction * seg_e + DB.XYZ(0, 0, pz - base_z)
                self._place(
                    place_beam(self.doc, s, e, plate_sym, level, PLATE_ROT),
                    "BOTTOM_PLATE",
                )

        # Mid-plate elevations are tier breaks for the studs. The mid-plate
        # members are placed with the infill stage after openings are framed.
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

        # ==== 1B. WALL SHAPE - TOP PLATES ====
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
                self._place(inst, "TOP_PLATE")
                # Track the lowest top-plate run on sloped walls for stud attachment.
                # A wall can have short flat segments even when the overall top is angled.
                if inst and i == 0 and wall_is_sloped:
                    seg_is_angled = abs(z_end - z_start) > (1.0 / 96.0)
                    _plate_beams.append((inst, seg_s, seg_e, seg_is_angled))

        # ==== 1C. WALL SHAPE - SIDE STUDS ====
        self._place_wall_side_studs(
            p0,
            direction,
            length,
            stud_bottom,
            top_profile,
            mid_plate_zs,
            stud_sym,
            level,
            wall_angle,
            base_z,
            occupied,
            _stud_columns,
        )

        # ==== 2A. OPENINGS - KING STUDS & JACK STUDS ====
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
                                             sl, "KING_STUD")
                    occupied.add(round(kd_c, 4))

            if self.config.include_jack_studs and head_z > stud_bottom:
                jack_h = head_z - stud_bottom
                for jd in (jack_left, jack_right):
                    jd_c = max(0.0, min(length, jd))
                    if self._near(jd_c, occupied):
                        continue
                    self._place_stud(p0, direction, jd_c, stud_bottom,
                                     jack_h, stud_sym, level, wall_angle, base_z,
                                     None, "JACK_STUD")
                    occupied.add(round(jd_c, 4))

        # ==== 2B. OPENINGS - HEADERS + SILLS + CRIPPLES ====
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
            span_s = max(0.0, span_s)
            span_e = min(length, span_e)
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
                    self._place(
                        place_beam(self.doc, hps, hpe, plate_sym, level, PLATE_ROT),
                        "HEADER_PLATE",
                    )

                # Header member(s) on-edge above head plate — side by side
                header_base = head_z + PLATE_THICKNESS
                hz = header_base + h_depth / 2.0
                for hi in range(h_count):
                    # Offset each header in wall-normal direction (face-to-face)
                    y_off = (hi - (h_count - 1) / 2.0) * STUD_THICKNESS
                    n_shift = normal * y_off
                    hs = p0 + direction * span_s + DB.XYZ(n_shift.X, n_shift.Y, hz - base_z)
                    he = p0 + direction * span_e + DB.XYZ(n_shift.X, n_shift.Y, hz - base_z)
                    self._place(
                        place_beam(self.doc, hs, he, header_sym, level, 0.0),
                        "HEADER",
                    )

                # Top plate above header assembly
                header_top_z = header_base + h_depth
                if plate_sym is not None:
                    tp_z = header_top_z + PLATE_THICKNESS / 2.0
                    tps = p0 + direction * span_s + DB.XYZ(0, 0, tp_z - base_z)
                    tpe = p0 + direction * span_e + DB.XYZ(0, 0, tp_z - base_z)
                    self._place(
                        place_beam(self.doc, tps, tpe, plate_sym, level, PLATE_ROT),
                        "HEADER_PLATE",
                    )
                    header_top_z += PLATE_THICKNESS

                # Cripple studs above header — aligned to wall OC grid
                mid_d = (left_e + right_e) / 2.0
                local_stud_top = self._top_of_stud_at(top_profile, mid_d)
                if self.config.include_cripple_studs and header_top_z < local_stud_top:
                    crip_h = local_stud_top - header_top_z
                    self._place_cripples_oc(p0, direction, left_e, right_e,
                                            header_top_z, crip_h, spacing,
                                            stud_sym, level, wall_angle, base_z,
                                            occupied, None, "CRIPPLE_STUD")

            # Sill plate for windows
            if is_win and sill_z > stud_bottom and plate_sym:
                sill_center = sill_z - PLATE_THICKNESS / 2.0
                if right_e - left_e >= MIN_LENGTH:
                    ss = p0 + direction * left_e  + DB.XYZ(0, 0, sill_center - base_z)
                    se = p0 + direction * right_e + DB.XYZ(0, 0, sill_center - base_z)
                    self._place(
                        place_beam(self.doc, ss, se, plate_sym, level, PLATE_ROT),
                        "SILL_PLATE",
                    )

                # Cripple studs below sill — aligned to wall OC grid
                if self.config.include_cripple_studs:
                    crip_top = sill_z - PLATE_THICKNESS
                    crip_h = crip_top - stud_bottom
                    if crip_h >= MIN_LENGTH:
                        self._place_cripples_oc(p0, direction, left_e, right_e,
                                                stud_bottom, crip_h, spacing,
                                                stud_sym, level, wall_angle, base_z,
                                                occupied, None, "CRIPPLE_STUD")

        # ==== 3A. INFILL - MID PLATES ====
        all_opening_gaps = [(op["left"], op["right"]) for op in openings]
        mid_plate_segs = self._split_segments(0.0, length, all_opening_gaps)

        for mz in mid_plate_zs:
            mid_center = mz + PLATE_THICKNESS / 2.0
            for seg_s, seg_e in mid_plate_segs:
                s = p0 + direction * seg_s + DB.XYZ(0, 0, mid_center - base_z)
                e = p0 + direction * seg_e + DB.XYZ(0, 0, mid_center - base_z)
                self._place(
                    place_beam(self.doc, s, e, plate_sym, level, PLATE_ROT),
                    "MID_PLATE",
                )

        # ==== 3B. INFILL - OC STUDS ====
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

        # ==== 3C. INFILL - MID-HEIGHT BLOCKING ====
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
                self._place(
                    place_beam(self.doc, bs, be, plate_sym, level, PLATE_ROT),
                    "BLOCKING",
                )

        # ==== FINALIZE. ATTACH STUD TOPS TO SLOPED PLATES ====
        self._attach_studs_to_plates(_stud_columns, _plate_beams)

        return True

    def _place_wall_side_studs(self, p0, direction, length, stud_bottom, top_profile,
                               mid_plate_zs, stud_sym, level, wall_angle, base_z,
                               occupied, stud_list):
        """Frame the actual wall boundary first, independent of wall joins."""
        if length < MIN_LENGTH:
            return

        edge_center = min(STUD_THICKNESS * 0.5, max(0.0, length * 0.5))
        positions = [edge_center]
        far_edge = max(0.0, length - edge_center)
        if abs(far_edge - positions[0]) > (STUD_THICKNESS * 0.5):
            positions.append(far_edge)

        for pos in positions:
            self._place_tiered_full_height_stud(
                p0,
                direction,
                pos,
                stud_bottom,
                top_profile,
                mid_plate_zs,
                stud_sym,
                level,
                wall_angle,
                base_z,
                occupied,
                stud_list,
                "SIDE_STUD",
            )

    def _place_tiered_full_height_stud(self, p0, direction, d, stud_bottom, top_profile,
                                       mid_plate_zs, stud_sym, level, wall_angle,
                                       base_z, occupied=None, stud_list=None,
                                       member_role="STUD"):
        """Place a full-height stud split at mid-plate tiers."""
        if occupied is not None and self._near(d, occupied, STUD_THICKNESS * 0.9):
            return False

        local_top = self._top_of_stud_at(top_profile, d)
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
                sl = stud_list if ti == last_tier else None
                self._place_stud(
                    p0,
                    direction,
                    d,
                    t_bot,
                    t_h,
                    stud_sym,
                    level,
                    wall_angle,
                    base_z,
                    sl,
                    member_role,
                )
                placed_any = True

        if placed_any and occupied is not None:
            occupied.add(round(d, 4))
        return placed_any

    def _track_member(self, inst, member_role):
        """Tag legacy wall placements so schedule backfill can resolve the host."""
        if inst is None or self._active_host_info is None:
            return
        try:
            tag_instance(
                inst,
                self._active_host_info,
                _WallTrackedMember(self._active_host_info, member_role),
            )
        except Exception:
            pass

    def _stamp_bom(self, inst, member_role, length_ft=None):
        if inst is None or self._active_host_info is None:
            return
        try:
            apply_bom_metadata(inst, self._active_host_info, member_role, length_ft)
        except Exception:
            pass

    def _place(self, inst, member_role=None):
        if inst:
            self.placed.append(inst)
            if member_role:
                self._track_member(inst, member_role)
                self._stamp_bom(inst, member_role)

    def _place_stud(self, p0, direction, d, bottom_z, height,
                    stud_sym, level, wall_angle, base_z,
                    stud_list=None, member_role="STUD"):
        pt = p0 + direction * d + DB.XYZ(0, 0, bottom_z - base_z)
        inst = place_column(self.doc, pt, height, stud_sym, level, wall_angle)
        if inst:
            self.placed.append(inst)
            self._track_member(inst, member_role)
            self._stamp_bom(inst, member_role, height)
            if stud_list is not None:
                stud_list.append((inst, d))

    def _place_cripples_oc(self, p0, direction, left, right, bottom_z,
                           height, oc_spacing, stud_sym, level, wall_angle,
                           base_z, occupied, stud_list=None,
                           member_role="CRIPPLE_STUD"):
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
                                     stud_list, member_role)
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
        wall_info = getattr(host_info, "wall_info", None)
        if wall_info is not None:
            base_z = wall_info.level_elevation + wall_info.base_offset
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

    def _wall_face_shape(self, wall, p0, direction, normal, base_z):
        """Read the actual wall side-face outline and openings from Revit."""
        best = None
        for face in self._get_wall_side_faces(wall, normal):
            loops = self._face_loops_local(face, p0, direction)
            if not loops:
                continue

            outer_index = self._largest_loop_index(loops)
            if outer_index is None:
                continue
            outer_loop = loops[outer_index]
            if len(outer_loop) < 3:
                continue

            start_d = min(d for d, _ in outer_loop)
            end_d = max(d for d, _ in outer_loop)
            span = end_d - start_d
            if span < MIN_LENGTH:
                continue

            shifted_outer = [(d - start_d, z) for d, z in outer_loop]
            top_profile = self._profile_from_loop_top(shifted_outer, span)
            if not top_profile or len(top_profile) < 2:
                continue

            openings = []
            for index, loop in enumerate(loops):
                if index == outer_index:
                    continue
                opening = self._opening_from_face_loop(loop, start_d, span, base_z)
                if opening is not None:
                    openings.append(opening)
            openings.sort(key=lambda item: item["left"])

            area = abs(self._loop_area_2d(outer_loop))
            if best is None or area > best["area"]:
                best = {
                    "area": area,
                    "start_d": start_d,
                    "length": span,
                    "top_profile": top_profile,
                    "openings": openings,
                }

        return best

    def _get_wall_side_faces(self, wall, wall_normal):
        """Get the major interior wall side faces."""
        faces = []
        seen = set()

        shell_layer = getattr(DB.ShellLayerType, "Interior", None)
        if shell_layer is not None:
            try:
                refs = DB.HostObjectUtils.GetSideFaces(wall, shell_layer)
            except Exception:
                refs = []
            for reference in refs:
                face = self._face_from_reference(wall, reference)
                if face is None:
                    continue
                self._add_unique_face(faces, seen, face)

        if faces:
            return faces

        try:
            opts = DB.Options()
            opts.ComputeReferences = True
            opts.DetailLevel = DB.ViewDetailLevel.Fine
            geom = wall.get_Geometry(opts)
        except Exception:
            geom = None

        if geom is None:
            return faces

        for geom_obj in geom:
            solids = []
            if hasattr(geom_obj, "Faces"):
                solids.append(geom_obj)
            elif hasattr(geom_obj, "GetInstanceGeometry"):
                try:
                    inst_geom = geom_obj.GetInstanceGeometry()
                except Exception:
                    inst_geom = None
                if inst_geom:
                    for inst_obj in inst_geom:
                        if hasattr(inst_obj, "Faces"):
                            solids.append(inst_obj)

            for solid in solids:
                for face in solid.Faces:
                    face_normal = self._face_normal(face)
                    if face_normal is None:
                        continue
                    if abs(face_normal.Z) > 0.1:
                        continue
                    dot = face_normal.DotProduct(wall_normal)
                    if dot > -0.8:
                        continue
                    self._add_unique_face(faces, seen, face)

        return faces

    @staticmethod
    def _face_from_reference(wall, reference):
        try:
            return wall.GetGeometryObjectFromReference(reference)
        except Exception:
            return None

    @staticmethod
    def _add_unique_face(faces, seen, face):
        face_id = getattr(face, "Id", None)
        if face_id is not None:
            try:
                marker = getattr(face_id, "IntegerValue", getattr(face_id, "Value", face_id))
            except Exception:
                marker = face_id
        else:
            marker = id(face)
        if marker in seen:
            return False
        seen.add(marker)
        faces.append(face)
        return True

    @staticmethod
    def _face_loops_local(face, p0, direction):
        """Project each face loop to local wall coordinates (distance, Z)."""
        loops = []
        try:
            curve_loops = face.GetEdgesAsCurveLoops()
        except Exception:
            curve_loops = None
        if curve_loops is None:
            return loops

        for curve_loop in curve_loops:
            points = []
            for curve in curve_loop:
                try:
                    tess_points = curve.Tessellate()
                except Exception:
                    tess_points = None
                if not tess_points:
                    continue
                for pt in tess_points:
                    try:
                        d = (pt - p0).DotProduct(direction)
                    except Exception:
                        continue
                    candidate = (d, pt.Z)
                    if points:
                        prev = points[-1]
                        if abs(prev[0] - candidate[0]) < 1e-6 and abs(prev[1] - candidate[1]) < 1e-6:
                            continue
                    points.append(candidate)
            if len(points) >= 3:
                loops.append(points)
        return loops

    @staticmethod
    def _largest_loop_index(loops):
        if not loops:
            return None
        best_index = None
        best_area = None
        for index, loop in enumerate(loops):
            area = abs(WallFramer._loop_area_2d(loop))
            if best_area is None or area > best_area:
                best_area = area
                best_index = index
        return best_index

    def _profile_from_loop_top(self, loop_points, span):
        """Build a top profile from a face loop's outer boundary."""
        if not loop_points:
            return None
        bucket_sz = 1.0 / 12.0
        buckets = {}
        for d, z in loop_points:
            if d < -bucket_sz or d > span + bucket_sz:
                continue
            clamped_d = max(0.0, min(span, d))
            key = round(clamped_d / bucket_sz)
            if key not in buckets or z > buckets[key][1]:
                buckets[key] = (clamped_d, z)
        if not buckets:
            return None
        profile = sorted(buckets.values())
        if profile[0][0] > 1e-6:
            profile.insert(0, (0.0, profile[0][1]))
        if profile[-1][0] < span - 1e-6:
            profile.append((span, profile[-1][1]))
        profile = self._simplify_profile(profile)
        return profile if len(profile) >= 2 else None

    def _opening_from_face_loop(self, loop_points, start_d, span, base_z):
        """Convert an inner face loop into an opening record."""
        if not loop_points:
            return None
        left = min(d for d, _ in loop_points) - start_d
        right = max(d for d, _ in loop_points) - start_d
        sill_z = min(z for _, z in loop_points)
        head_z = max(z for _, z in loop_points)

        left = max(0.0, left)
        right = min(span, right)
        edge_tol = STUD_THICKNESS * 2.0
        if left <= edge_tol or right >= span - edge_tol:
            return None
        if right - left < 0.5 or head_z - sill_z < 0.5:
            return None
        return {
            "left": left,
            "right": right,
            "head_z": head_z,
            "sill_z": sill_z,
            "is_window": sill_z > base_z + PLATE_THICKNESS,
        }

    @staticmethod
    def _merge_opening_sets(primary, secondary):
        """Merge two opening lists while removing near-duplicate spans."""
        merged = []
        for source in (primary or [], secondary or []):
            for opening in source:
                is_dup = False
                for existing in merged:
                    if (abs(existing["left"] - opening["left"]) < STUD_THICKNESS and
                            abs(existing["right"] - opening["right"]) < STUD_THICKNESS and
                            abs(existing["head_z"] - opening["head_z"]) < STUD_THICKNESS):
                        is_dup = True
                        break
                if not is_dup:
                    merged.append(dict(opening))
        merged.sort(key=lambda item: item["left"])
        return merged

    @staticmethod
    def _loop_area_2d(loop_points):
        if not loop_points:
            return 0.0
        points = list(loop_points)
        if len(points) < 3:
            return 0.0
        area = 0.0
        for index in range(len(points)):
            x0, y0 = points[index]
            x1, y1 = points[(index + 1) % len(points)]
            area += (x0 * y1) - (x1 * y0)
        return area * 0.5

    @staticmethod
    def _face_normal(face):
        try:
            bbox = face.GetBoundingBox()
            if bbox is None:
                return None
            uv = DB.UV(
                (bbox.Min.U + bbox.Max.U) * 0.5,
                (bbox.Min.V + bbox.Max.V) * 0.5,
            )
            normal = face.ComputeNormal(uv)
            if normal is None:
                return None
            return normal.Normalize()
        except Exception:
            return None

    @staticmethod
    def _trim_openings_to_span(openings, start_d, end_d):
        """Shift/clamp opening distances into a trimmed wall run."""
        span = max(0.0, end_d - start_d)
        trimmed = []
        for opening in openings or []:
            left = opening["left"] - start_d
            right = opening["right"] - start_d
            if right <= 0.0 or left >= span:
                continue

            clipped = dict(opening)
            clipped["left"] = max(0.0, left)
            clipped["right"] = min(span, right)
            if clipped["right"] - clipped["left"] < MIN_LENGTH:
                continue
            trimmed.append(clipped)

        trimmed.sort(key=lambda item: item["left"])
        return trimmed

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

        cfg.wall_layer_mode = LAYER_MODE_STRUCTURAL

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


def _support_top_elevation(element):
    """Return a coarse top elevation fallback for wall base alignment."""
    if element is None:
        return None
    try:
        bbox = element.get_BoundingBox(None)
        if bbox is not None:
            return bbox.Max.Z
    except Exception:
        pass
    return None


def _delete_existing_wall_framing(doc, walls):
    """Delete tracked framing previously generated for the selected walls."""
    wall_ids = set()
    for wall in walls or []:
        wall_id = _element_id_text(getattr(wall, "Id", None))
        if wall_id:
            wall_ids.add(wall_id)
    if not wall_ids:
        return 0

    delete_ids = []
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
            try:
                tracking = get_tracking_data(element)
            except Exception:
                tracking = None
            if not tracking:
                continue
            if tracking.get("kind") != "wall":
                continue
            if tracking.get("host") not in wall_ids:
                continue
            delete_ids.append(element.Id)

    deleted = 0
    for element_id in delete_ids:
        try:
            doc.Delete(element_id)
            deleted += 1
        except Exception:
            pass
    return deleted


def _element_id_text(element_id):
    if element_id is None:
        return None
    value = getattr(element_id, "IntegerValue", getattr(element_id, "Value", None))
    if value is None:
        return None
    return str(value)


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
        config.wall_base_override_z = _support_top_elevation(support)

    # Persist user options but not transient support element references.
    persist_cfg = FramingConfig.from_dict(config.to_dict())
    persist_cfg.wall_base_support_element_id = None
    persist_cfg.wall_base_override_z = None
    _save_last_config(persist_cfg)

    framer = WallFramer(doc, config)
    total = 0
    deleted_existing = 0
    with revit.Transaction("WF: Frame Walls"):
        try:
            ensure_bom_parameters(doc)
        except Exception as bom_param_err:
            logger.warning(
                "BOM parameter setup skipped during wall framing: {0}".format(
                    bom_param_err
                )
            )
        deleted_existing = _delete_existing_wall_framing(doc, walls)
        for wall in walls:
            if framer.frame_wall(wall):
                total += 1

    output.print_md(
        "## Wood Framing Complete\n"
        "- **Engine:** {0}\n"
        "- **Walls framed:** {1}\n"
        "- **Existing tracked members deleted:** {2}\n"
        "- **Members placed:** {3}\n".format(
            FRAME_WALL_ENGINE,
            total,
            deleted_existing,
            len(framer.placed),
        )
    )


if __name__ == "__main__":
    main()
