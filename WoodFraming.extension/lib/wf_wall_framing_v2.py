# -*- coding: utf-8 -*-
"""Wall Framing 2.0 - isolated face-based wall framing engine.

This module does not call the legacy wall framing engine. It builds a local
wall model from the interior wall side face, chooses the structural/core layer
for placement, then emits members in three stages:
wall shape, openings, infill.
"""

import math

from wf_config import LAYER_MODE_STRUCTURAL, WALL_BASE_MODE_SUPPORT_TOP
from wf_geometry import (
    FramingMember,
    _get_opening_height,
    _get_opening_width,
    _get_sill_height,
    inches_to_feet,
    safe_wall_normal,
)
from wf_host import (
    _build_compound_layers,
    _get_compound_structure,
    _preferred_wall_target_layer,
    _select_target_layer,
)
from wf_placement import BaseFramingEngine


ENGINE_NAME = "wall-framing-2.0-isolated"
MIN_MEMBER_LENGTH = inches_to_feet(1.0)
PLATE_THICKNESS = inches_to_feet(1.5)
STUD_THICKNESS = inches_to_feet(1.5)
PLATE_ROTATION = -math.pi / 2.0
HEADER_ROTATION = 0.0
MID_PLATE_INTERVAL = 8.0


class WallFaceOpeningV2(object):
    def __init__(self, left, right, sill_height, head_height, is_window):
        self.left_edge = max(0.0, left)
        self.right_edge = max(self.left_edge, right)
        self.sill_height = max(0.0, sill_height)
        self.head_height = max(self.sill_height, head_height)
        self.is_window = bool(is_window)
        self.is_door = not self.is_window
        self.distance_along_wall = (self.left_edge + self.right_edge) * 0.5


class WallFaceHostInfoV2(object):
    def __init__(self):
        self.kind = "wall_v2"
        self.element = None
        self.element_id = None
        self.level_id = None
        self.level_elevation = 0.0
        self.base_elevation = 0.0
        self.start_point = None
        self.direction = None
        self.normal = None
        self.length = 0.0
        self.angle = 0.0
        self.target_layer = None
        self.top_profile = []
        self.openings = []
        self.audit = {}

    def point_at(self, distance_along, height, lateral_offset=0.0):
        from Autodesk.Revit.DB import XYZ

        point = self.start_point + self.direction * distance_along
        if abs(lateral_offset) > 1e-9:
            point = point + self.normal * lateral_offset
        return XYZ(point.X, point.Y, self.base_elevation + height)

    def top_abs_at(self, distance_along):
        return _interpolate_profile(self.top_profile, distance_along)

    def height_at(self, distance_along):
        return self.top_abs_at(distance_along) - self.base_elevation


class WallFaceFramingV2Engine(BaseFramingEngine):
    """Calculate and place wall framing from an interior side-face outline."""

    def calculate_members(self, wall):
        host = self._analyze_wall(wall)
        if host is None:
            return [], None

        occupied = set()
        members = []

        members.extend(self._wall_shape_members(host, occupied))
        members.extend(self._opening_members(host, occupied))
        members.extend(self._infill_members(host, occupied))

        return members, host

    # ------------------------------------------------------------------
    # Analysis
    # ------------------------------------------------------------------

    def _analyze_wall(self, wall):
        from Autodesk.Revit.DB import BuiltInParameter, Line, WallLocationLine, XYZ

        loc = getattr(wall, "Location", None)
        if loc is None:
            return None
        curve = loc.Curve
        if not isinstance(curve, Line):
            return None

        p0 = curve.GetEndPoint(0)
        p1 = curve.GetEndPoint(1)
        direction = (p1 - p0).Normalize()
        normal = safe_wall_normal(wall, direction)
        if normal is None:
            return None

        level = self.doc.GetElement(wall.LevelId)
        if level is None:
            return None

        base_z = level.Elevation
        base_offset = wall.get_Parameter(BuiltInParameter.WALL_BASE_OFFSET)
        if base_offset is not None and base_offset.HasValue:
            base_z += base_offset.AsDouble()
        if getattr(self.config, "wall_base_mode", None) == WALL_BASE_MODE_SUPPORT_TOP:
            override = getattr(self.config, "wall_base_override_z", None)
            if override is not None:
                try:
                    base_z = float(override)
                except Exception:
                    pass

        compound = _get_compound_structure(wall)
        layers = _build_compound_layers(self.doc, compound)
        target_layer = _select_target_layer(layers, LAYER_MODE_STRUCTURAL)
        target_layer = _preferred_wall_target_layer(
            layers,
            LAYER_MODE_STRUCTURAL,
            target_layer,
        )
        target_offset = self._target_offset_from_location_line(
            wall,
            compound,
            target_layer,
            WallLocationLine,
        )

        target_origin = p0 + normal * target_offset
        target_origin = XYZ(target_origin.X, target_origin.Y, base_z)

        face_shape = self._interior_face_shape(wall, target_origin, direction, base_z)
        if face_shape is None:
            return None

        host = WallFaceHostInfoV2()
        host.element = wall
        host.element_id = wall.Id
        host.level_id = wall.LevelId
        host.level_elevation = level.Elevation
        host.base_elevation = base_z
        host.start_point = target_origin + direction * face_shape["start_d"]
        host.direction = direction
        host.normal = normal
        host.length = face_shape["length"]
        host.angle = math.atan2(direction.Y, direction.X)
        host.target_layer = target_layer
        host.top_profile = face_shape["top_profile"]
        host.audit = {
            "wall_id": _element_id_text(wall.Id),
            "location_length": curve.Length,
            "face_start_shift": face_shape["start_d"],
            "face_length": face_shape["length"],
            "face_loop_count": face_shape.get("loop_count", 0),
            "face_opening_count": len(face_shape.get("openings", [])),
            "target_offset": target_offset,
            "target_layer_index": getattr(target_layer, "index", None),
            "target_layer_function": getattr(target_layer, "function", None),
            "target_layer_width": getattr(target_layer, "width", None),
        }

        hosted_openings = self._hosted_openings(
            wall,
            target_origin,
            direction,
            face_shape["start_d"],
            face_shape["start_d"] + face_shape["length"],
            base_z,
        )
        host.openings = self._merge_openings(
            face_shape["openings"],
            hosted_openings,
            host.length,
            base_z,
        )
        host.audit["hosted_opening_count"] = len(hosted_openings)
        host.audit["merged_opening_count"] = len(host.openings)
        return host

    @staticmethod
    def _target_offset_from_location_line(wall, compound, target_layer,
                                          wall_location_line_type):
        if compound is None or target_layer is None:
            return 0.0
        current_line = _wall_location_line(wall, wall_location_line_type)
        try:
            current_offset = compound.GetOffsetForLocationLine(current_line)
            return target_layer.center_offset - current_offset
        except Exception:
            return 0.0

    def _interior_face_shape(self, wall, origin, direction, base_z):
        face = self._interior_side_face(wall)
        if face is None:
            return None
        loops = self._face_loops_local(face, origin, direction)
        if not loops:
            return None

        outer_index = self._largest_loop_index(loops)
        if outer_index is None:
            return None
        outer_loop = loops[outer_index]
        if len(outer_loop) < 3:
            return None

        start_d = min(d for d, _ in outer_loop)
        end_d = max(d for d, _ in outer_loop)
        span = end_d - start_d
        if span < MIN_MEMBER_LENGTH:
            return None

        shifted_outer = [(d - start_d, z) for d, z in outer_loop]
        top_profile = self._top_profile_from_loop(shifted_outer, span)
        if not top_profile:
            return None

        openings = []
        for index, loop in enumerate(loops):
            if index == outer_index:
                continue
            opening = self._opening_from_loop(loop, start_d, span, base_z)
            if opening is not None:
                openings.append(opening)

        return {
            "start_d": start_d,
            "length": span,
            "top_profile": top_profile,
            "openings": openings,
            "loop_count": len(loops),
        }

    @staticmethod
    def _interior_side_face(wall):
        from Autodesk.Revit.DB import HostObjectUtils, ShellLayerType

        try:
            refs = HostObjectUtils.GetSideFaces(wall, ShellLayerType.Interior)
        except Exception:
            refs = []
        best_face = None
        best_area = 0.0
        for reference in refs:
            try:
                face = wall.GetGeometryObjectFromReference(reference)
            except Exception:
                face = None
            if face is None:
                continue
            area = float(getattr(face, "Area", 0.0) or 0.0)
            if best_face is None or area > best_area:
                best_face = face
                best_area = area
        return best_face

    @staticmethod
    def _face_loops_local(face, origin, direction):
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
                    d = (pt - origin).DotProduct(direction)
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
        best_index = None
        best_area = None
        for index, loop in enumerate(loops):
            area = abs(_polygon_area(loop))
            if best_area is None or area > best_area:
                best_area = area
                best_index = index
        return best_index

    def _top_profile_from_loop(self, loop_points, span):
        bucket = inches_to_feet(1.0)
        buckets = {}
        for d, z in loop_points:
            if d < -bucket or d > span + bucket:
                continue
            local_d = max(0.0, min(span, d))
            key = round(local_d / bucket)
            if key not in buckets or z > buckets[key][1]:
                buckets[key] = (local_d, z)
        if not buckets:
            return None
        profile = sorted(buckets.values())
        if profile[0][0] > 1e-6:
            profile.insert(0, (0.0, profile[0][1]))
        if profile[-1][0] < span - 1e-6:
            profile.append((span, profile[-1][1]))
        return _simplify_profile(profile)

    @staticmethod
    def _opening_from_loop(loop_points, start_d, span, base_z):
        left = min(d for d, _ in loop_points) - start_d
        right = max(d for d, _ in loop_points) - start_d
        sill_abs = min(z for _, z in loop_points)
        head_abs = max(z for _, z in loop_points)
        left = max(0.0, left)
        right = min(span, right)
        edge_tol = STUD_THICKNESS * 2.0
        if left <= edge_tol or right >= span - edge_tol:
            return None
        if right - left < inches_to_feet(6.0):
            return None
        if head_abs - sill_abs < inches_to_feet(6.0):
            return None
        return WallFaceOpeningV2(
            left,
            right,
            sill_abs - base_z,
            head_abs - base_z,
            sill_abs > base_z + PLATE_THICKNESS * 2.0,
        )

    def _hosted_openings(self, wall, origin, direction, start_d, end_d, base_z):
        openings = []
        try:
            insert_ids = wall.FindInserts(True, False, False, False)
        except Exception:
            insert_ids = []

        for insert_id in insert_ids:
            try:
                elem = self.doc.GetElement(insert_id)
            except Exception:
                elem = None
            if elem is None:
                continue

            opening = self._opening_from_insert(elem, origin, direction, start_d, end_d, base_z)
            if opening is not None:
                openings.append(opening)

        return openings

    def _opening_from_insert(self, elem, origin, direction, start_d, end_d, base_z):
        from Autodesk.Revit.DB import BuiltInCategory, Opening

        try:
            if isinstance(elem, Opening):
                boundary = elem.BoundaryRect
                if boundary and len(boundary) >= 2:
                    d0 = (boundary[0] - origin).DotProduct(direction)
                    d1 = (boundary[1] - origin).DotProduct(direction)
                    z0 = boundary[0].Z
                    z1 = boundary[1].Z
                    return self._opening_from_abs_span(
                        min(d0, d1),
                        max(d0, d1),
                        min(z0, z1),
                        max(z0, z1),
                        start_d,
                        end_d,
                        base_z,
                    )
        except Exception:
            pass

        category = getattr(elem, "Category", None)
        category_id = None
        if category is not None:
            category_id = getattr(category.Id, "IntegerValue", getattr(category.Id, "Value", None))

        is_window = category_id == int(BuiltInCategory.OST_Windows)
        is_door = category_id == int(BuiltInCategory.OST_Doors)

        if is_window or is_door:
            loc = getattr(elem, "Location", None)
            point = getattr(loc, "Point", None)
            if point is None:
                return None
            center_d = (point - origin).DotProduct(direction)
            width = self._read_opening_dimension(elem, "width")
            height = self._read_opening_dimension(elem, "height")
            sill = self._read_sill_height(elem) if is_window else 0.0
            return self._opening_from_abs_span(
                center_d - width * 0.5,
                center_d + width * 0.5,
                base_z + sill,
                base_z + sill + height,
                start_d,
                end_d,
                base_z,
            )

        return None

    @staticmethod
    def _opening_from_abs_span(left_abs_d, right_abs_d, sill_abs, head_abs,
                               start_d, end_d, base_z):
        left = left_abs_d - start_d
        right = right_abs_d - start_d
        span = max(0.0, end_d - start_d)
        if right <= 0.0 or left >= span:
            return None
        left = max(0.0, left)
        right = min(span, right)
        if right - left < inches_to_feet(6.0):
            return None
        if head_abs - sill_abs < inches_to_feet(6.0):
            return None
        return WallFaceOpeningV2(
            left,
            right,
            sill_abs - base_z,
            head_abs - base_z,
            sill_abs > base_z + PLATE_THICKNESS * 2.0,
        )

    @staticmethod
    def _read_opening_dimension(elem, dim_kind):
        if dim_kind == "width":
            return _get_opening_width(elem)
        return _get_opening_height(elem)

    @staticmethod
    def _read_sill_height(elem):
        return _get_sill_height(elem)

    @staticmethod
    def _merge_openings(face_openings, hosted_openings, length, base_z):
        merged = []
        for source in (hosted_openings or [], face_openings or []):
            for opening in source:
                if opening.right_edge <= 0.0 or opening.left_edge >= length:
                    continue
                left = max(0.0, opening.left_edge)
                right = min(length, opening.right_edge)
                if right - left < inches_to_feet(6.0):
                    continue
                candidate = WallFaceOpeningV2(
                    left,
                    right,
                    opening.sill_height,
                    opening.head_height,
                    opening.sill_height > PLATE_THICKNESS * 2.0,
                )
                duplicate = False
                for existing in merged:
                    if (abs(existing.left_edge - candidate.left_edge) < STUD_THICKNESS and
                            abs(existing.right_edge - candidate.right_edge) < STUD_THICKNESS):
                        duplicate = True
                        break
                if not duplicate:
                    merged.append(candidate)
        merged.sort(key=lambda op: op.left_edge)
        return merged

    # ------------------------------------------------------------------
    # Member calculation
    # ------------------------------------------------------------------

    def _wall_shape_members(self, host, occupied):
        members = []
        members.extend(self._bottom_plates(host))
        members.extend(self._top_plates(host))
        members.extend(self._side_studs(host, occupied))
        return members

    def _bottom_plates(self, host):
        members = []
        family = self.config.bottom_plate_family_name or self.config.stud_family_name
        type_name = self.config.bottom_plate_type_name or self.config.stud_type_name
        door_gaps = []
        for op in host.openings:
            if not op.is_window:
                door_gaps.append((op.left_edge, op.right_edge))
        segments = _split_segments(0.0, host.length, door_gaps)

        for index in range(int(getattr(self.config, "bottom_plate_count", 1))):
            center_h = index * PLATE_THICKNESS + PLATE_THICKNESS * 0.5
            for start_d, end_d in segments:
                members.append(
                    self._beam(
                        "BOTTOM_PLATE",
                        host,
                        start_d,
                        end_d,
                        center_h,
                        center_h,
                        family,
                        type_name,
                        PLATE_ROTATION,
                    )
                )
        return [m for m in members if m is not None]

    def _top_plates(self, host):
        members = []
        family = self.config.top_plate_family_name or self.config.bottom_plate_family_name or self.config.stud_family_name
        type_name = self.config.top_plate_type_name or self.config.bottom_plate_type_name or self.config.stud_type_name
        top_count = int(getattr(self.config, "top_plate_count", 2))
        profile = _simplify_profile(host.top_profile)
        for profile_index in range(len(profile) - 1):
            d0, z0 = profile[profile_index]
            d1, z1 = profile[profile_index + 1]
            if d1 - d0 < MIN_MEMBER_LENGTH:
                continue
            for plate_index in range(top_count):
                off = (top_count - plate_index - 1) * PLATE_THICKNESS
                off += PLATE_THICKNESS * 0.5
                start_h = z0 - off - host.base_elevation
                end_h = z1 - off - host.base_elevation
                members.append(
                    self._beam(
                        "TOP_PLATE",
                        host,
                        d0,
                        d1,
                        start_h,
                        end_h,
                        family,
                        type_name,
                        PLATE_ROTATION,
                    )
                )
        return [m for m in members if m is not None]

    def _side_studs(self, host, occupied):
        members = []
        positions = [min(STUD_THICKNESS * 0.5, host.length * 0.5)]
        far = max(0.0, host.length - positions[0])
        if abs(far - positions[0]) > STUD_THICKNESS:
            positions.append(far)
        for pos in positions:
            member = self._full_height_stud("SIDE_STUD", host, pos)
            if member is not None:
                members.append(member)
                occupied.add(round(pos, 4))
        return members

    def _opening_members(self, host, occupied):
        members = []
        header_family = self.config.header_family_name or self.config.stud_family_name
        header_type = self.config.header_type_name or self.config.stud_type_name
        header_depth = self._header_depth(header_family, header_type)

        for op in host.openings:
            left = op.left_edge
            right = op.right_edge
            if right - left < MIN_MEMBER_LENGTH:
                continue

            if getattr(self.config, "include_king_studs", True):
                for pos in (left - STUD_THICKNESS * 1.5, right + STUD_THICKNESS * 1.5):
                    if pos <= 0.0 or pos >= host.length:
                        continue
                    if _near(pos, occupied, STUD_THICKNESS):
                        continue
                    member = self._full_height_stud("KING_STUD", host, pos)
                    if member is not None:
                        members.append(member)
                        occupied.add(round(pos, 4))

            if getattr(self.config, "include_jack_studs", True):
                for pos in (left - STUD_THICKNESS * 0.5, right + STUD_THICKNESS * 0.5):
                    if pos <= 0.0 or pos >= host.length:
                        continue
                    if _near(pos, occupied, STUD_THICKNESS):
                        continue
                    member = self._stud("JACK_STUD", host, pos, self._stud_bottom(), op.head_height)
                    if member is not None:
                        members.append(member)
                        occupied.add(round(pos, 4))

            span_start = max(0.0, left - STUD_THICKNESS)
            span_end = min(host.length, right + STUD_THICKNESS)
            if span_end - span_start >= MIN_MEMBER_LENGTH:
                header_center = op.head_height + header_depth * 0.5
                header_count = int(getattr(self.config, "header_count", 2))
                for header_index in range(header_count):
                    lateral = (header_index - (header_count - 1) * 0.5) * STUD_THICKNESS
                    members.append(
                        self._beam(
                            "HEADER",
                            host,
                            span_start,
                            span_end,
                            header_center,
                            header_center,
                            header_family,
                            header_type,
                            HEADER_ROTATION,
                            lateral,
                        )
                    )

            header_top = op.head_height + header_depth
            stud_top = self._stud_top(host, (left + right) * 0.5)
            if getattr(self.config, "include_cripple_studs", True) and header_top < stud_top:
                members.extend(self._cripples(host, left, right, header_top, stud_top, occupied))

            if op.is_window and op.sill_height > self._stud_bottom():
                members.append(
                    self._beam(
                        "SILL_PLATE",
                        host,
                        left,
                        right,
                        op.sill_height - PLATE_THICKNESS * 0.5,
                        op.sill_height - PLATE_THICKNESS * 0.5,
                        self.config.bottom_plate_family_name or self.config.stud_family_name,
                        self.config.bottom_plate_type_name or self.config.stud_type_name,
                        PLATE_ROTATION,
                    )
                )
                cripple_top = op.sill_height - PLATE_THICKNESS
                if getattr(self.config, "include_cripple_studs", True) and cripple_top > self._stud_bottom():
                    members.extend(self._cripples(host, left, right, self._stud_bottom(), cripple_top, occupied))

        return [m for m in members if m is not None]

    def _infill_members(self, host, occupied):
        members = []
        members.extend(self._mid_plates(host))
        members.extend(self._regular_studs(host, occupied))
        members.extend(self._blocking(host, occupied))
        return members

    def _mid_plates(self, host):
        members = []
        if not bool(getattr(self.config, "include_mid_plates", True)):
            return members
        interval = float(getattr(self.config, "mid_plate_interval_ft", MID_PLATE_INTERVAL))
        if interval <= 0.0:
            return members
        min_top = min(self._stud_top(host, 0.0), self._stud_top(host, host.length))
        z = self._stud_bottom() + interval
        gaps = [(op.left_edge, op.right_edge) for op in host.openings]
        segments = _split_segments(0.0, host.length, gaps)
        family = self.config.bottom_plate_family_name or self.config.stud_family_name
        type_name = self.config.bottom_plate_type_name or self.config.stud_type_name
        while z < min_top - PLATE_THICKNESS:
            center = z + PLATE_THICKNESS * 0.5
            for start_d, end_d in segments:
                members.append(
                    self._beam(
                        "MID_PLATE",
                        host,
                        start_d,
                        end_d,
                        center,
                        center,
                        family,
                        type_name,
                        PLATE_ROTATION,
                    )
                )
            z += interval
        return [m for m in members if m is not None]

    def _regular_studs(self, host, occupied):
        members = []
        spacing = self.config.stud_spacing_ft
        if spacing <= 0.0:
            return members
        d = spacing
        while d < host.length - STUD_THICKNESS * 0.5:
            if not self._in_opening_zone(host, d) and not _near(d, occupied, STUD_THICKNESS):
                member = self._full_height_stud("STUD", host, d)
                if member is not None:
                    members.append(member)
                    occupied.add(round(d, 4))
            d += spacing
        return members

    def _blocking(self, host, occupied):
        members = []
        spacing = self.config.stud_spacing_ft
        if spacing <= 0.0:
            return members
        sorted_pos = sorted(occupied)
        family = self.config.bottom_plate_family_name or self.config.stud_family_name
        type_name = self.config.bottom_plate_type_name or self.config.stud_type_name
        for index in range(len(sorted_pos) - 1):
            left = sorted_pos[index]
            right = sorted_pos[index + 1]
            if right - left < MIN_MEMBER_LENGTH or right - left > spacing * 1.5:
                continue
            mid = (left + right) * 0.5
            stud_height = self._stud_top(host, mid) - self._stud_bottom()
            if stud_height <= 8.0:
                continue
            z = self._stud_bottom() + stud_height * 0.5
            members.append(
                self._beam(
                    "BLOCKING",
                    host,
                    left + STUD_THICKNESS * 0.5,
                    right - STUD_THICKNESS * 0.5,
                    z,
                    z,
                    family,
                    type_name,
                    PLATE_ROTATION,
                )
            )
        return [m for m in members if m is not None]

    # ------------------------------------------------------------------
    # Member factories
    # ------------------------------------------------------------------

    def _beam(self, member_type, host, start_d, end_d, start_h, end_h,
              family, type_name, rotation, lateral=0.0):
        if end_d - start_d < MIN_MEMBER_LENGTH:
            return None
        start = host.point_at(start_d, start_h, lateral)
        end = host.point_at(end_d, end_h, lateral)
        if start.DistanceTo(end) < MIN_MEMBER_LENGTH:
            return None
        member = FramingMember(member_type, start, end)
        member.member_type = member_type
        member.family_name = family
        member.type_name = type_name
        self._tag(member, host, rotation, False)
        member.disallow_end_joins = True
        return member

    def _full_height_stud(self, member_type, host, d):
        return self._stud(member_type, host, d, self._stud_bottom(), self._stud_top(host, d))

    def _stud(self, member_type, host, d, bottom_h, top_h):
        if top_h - bottom_h < MIN_MEMBER_LENGTH:
            return None
        start = host.point_at(d, bottom_h)
        end = host.point_at(d, top_h)
        member = FramingMember(member_type, start, end)
        member.member_type = member_type
        member.family_name = self.config.stud_family_name
        member.type_name = self.config.stud_type_name
        self._tag(member, host, host.angle, True)
        return member

    def _cripples(self, host, left, right, bottom_h, top_h, occupied):
        members = []
        spacing = self.config.stud_spacing_ft
        if spacing <= 0.0 or top_h - bottom_h < MIN_MEMBER_LENGTH:
            return members
        d = spacing
        while d < right - STUD_THICKNESS * 0.5:
            if left + STUD_THICKNESS * 0.5 < d < right - STUD_THICKNESS * 0.5:
                if not _near(d, occupied, STUD_THICKNESS):
                    member = self._stud("CRIPPLE_STUD", host, d, bottom_h, top_h)
                    if member is not None:
                        members.append(member)
            d += spacing
        return members

    def _tag(self, member, host, rotation, is_column):
        member.rotation = rotation
        member.is_column = is_column
        member.host_kind = host.kind
        member.host_id = host.element_id
        if host.target_layer is not None:
            member.layer_index = host.target_layer.index

    def _stud_bottom(self):
        return PLATE_THICKNESS * int(getattr(self.config, "bottom_plate_count", 1))

    def _stud_top(self, host, d):
        return host.height_at(d) - PLATE_THICKNESS * int(getattr(self.config, "top_plate_count", 2))

    def _header_depth(self, family, type_name):
        depth = self.get_type_depth(family, type_name)
        if depth is not None and depth > 0.0:
            return depth
        return inches_to_feet(3.5)

    def _in_opening_zone(self, host, d):
        for op in host.openings:
            if op.left_edge - STUD_THICKNESS * 3.0 <= d <= op.right_edge + STUD_THICKNESS * 3.0:
                return True
        return False


def _wall_location_line(wall, wall_location_line_type):
    from Autodesk.Revit.DB import BuiltInParameter
    import System

    param = wall.get_Parameter(BuiltInParameter.WALL_KEY_REF_PARAM)
    if param is None or not param.HasValue:
        return wall_location_line_type.WallCenterline
    try:
        raw = param.AsInteger()
        if not System.Enum.IsDefined(wall_location_line_type, raw):
            return wall_location_line_type.WallCenterline
        return System.Enum.ToObject(wall_location_line_type, raw)
    except Exception:
        return wall_location_line_type.WallCenterline


def _element_id_text(element_id):
    if element_id is None:
        return None
    value = getattr(element_id, "IntegerValue", getattr(element_id, "Value", None))
    if value is None:
        return None
    return str(value)


def _polygon_area(points):
    area = 0.0
    for index in range(len(points)):
        x0, y0 = points[index]
        x1, y1 = points[(index + 1) % len(points)]
        area += (x0 * y1) - (x1 * y0)
    return area * 0.5


def _interpolate_profile(profile, d):
    if not profile:
        return 0.0
    if d <= profile[0][0]:
        return profile[0][1]
    if d >= profile[-1][0]:
        return profile[-1][1]
    for index in range(len(profile) - 1):
        d0, z0 = profile[index]
        d1, z1 = profile[index + 1]
        if d0 <= d <= d1:
            if abs(d1 - d0) < 1e-9:
                return z0
            t = (d - d0) / (d1 - d0)
            return z0 + (z1 - z0) * t
    return profile[-1][1]


def _simplify_profile(profile, tol=inches_to_feet(1.0)):
    if len(profile) <= 2:
        return list(profile)
    result = [profile[0]]
    for index in range(1, len(profile) - 1):
        dp, zp = result[-1]
        dc, zc = profile[index]
        dn, zn = profile[index + 1]
        if abs(dn - dp) < 1e-9:
            continue
        t = (dc - dp) / (dn - dp)
        expected = zp + (zn - zp) * t
        if abs(zc - expected) > tol:
            result.append(profile[index])
    result.append(profile[-1])
    return result


def _split_segments(start, end, gaps):
    segments = [(start, end)]
    for gap_start, gap_end in gaps:
        next_segments = []
        for seg_start, seg_end in segments:
            if gap_end <= seg_start or gap_start >= seg_end:
                next_segments.append((seg_start, seg_end))
                continue
            if gap_start > seg_start:
                next_segments.append((seg_start, gap_start))
            if gap_end < seg_end:
                next_segments.append((gap_end, seg_end))
        segments = next_segments
    return [(seg_start, seg_end) for seg_start, seg_end in segments
            if seg_end - seg_start >= MIN_MEMBER_LENGTH]


def _near(value, occupied, tol):
    for existing in occupied:
        if abs(value - existing) < tol:
            return True
    return False
