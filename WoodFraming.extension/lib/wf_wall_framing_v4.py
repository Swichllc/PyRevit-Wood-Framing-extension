# -*- coding: utf-8 -*-
"""Wall Framing 4.0 - wall-cavity driven framing engine.

This engine keeps the existing 2.0 command UI but replaces the framing model.
It uses the selected wall side face as the wall profile, builds scanline
intervals from the actual face loops, and validates candidate members against
the wall solid before they are handed to the shared placement service.
"""

import math

from wf_config import LAYER_MODE_STRUCTURAL, WALL_BASE_MODE_SUPPORT_TOP
from wf_families import find_family_symbol
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


ENGINE_NAME = "wall-framing-4.0-cavity"
MIN_MEMBER_LENGTH = inches_to_feet(1.0)
PLATE_THICKNESS = inches_to_feet(1.5)
STUD_THICKNESS = inches_to_feet(1.5)
DEFAULT_LUMBER_DEPTH = inches_to_feet(3.5)
DEFAULT_LUMBER_WIDTH = inches_to_feet(1.5)
DEFAULT_HEADER_DEPTH = inches_to_feet(3.5)
MID_PLATE_INTERVAL = 8.0
PLATE_ROTATION = -math.pi / 2.0
HEADER_ROTATION = 0.0
GEOM_TOL = inches_to_feet(0.125)
VALIDATION_TOL = inches_to_feet(0.25)
PROFILE_EDGE_TOL = STUD_THICKNESS * 2.0
OPENING_STUD_COLLISION_TOL = STUD_THICKNESS * 0.45
HEADER_PLY_SPACER = inches_to_feet(0.5)
MIN_HEADER_PLY_COUNT = 2
OPENING_FRAME_MEMBER_TYPES = set([
    "HEADER",
    "HEADER_BOTTOM_PLATE",
    "HEADER_TOP_PLATE",
    "JACK_STUD",
    "KING_STUD",
    "SILL_PLATE",
    "CRIPPLE_STUD",
])


class WallCavityOpeningV4(object):
    def __init__(self, left, right, sill_height, head_height, is_window, source):
        self.left_edge = max(0.0, float(left))
        self.right_edge = max(self.left_edge, float(right))
        self.sill_height = max(0.0, float(sill_height))
        self.head_height = max(self.sill_height, float(head_height))
        self.is_window = bool(is_window)
        self.is_door = not self.is_window
        self.source = source
        self.distance_along_wall = (self.left_edge + self.right_edge) * 0.5


class WallCavitySegmentV4(object):
    def __init__(self, kind, p0, p1, d0, d1, z0, z1):
        self.kind = kind
        self.p0 = p0
        self.p1 = p1
        self.d0 = float(d0)
        self.d1 = float(d1)
        self.z0 = float(z0)
        self.z1 = float(z1)

    def d_min(self):
        return min(self.d0, self.d1)

    def d_max(self):
        return max(self.d0, self.d1)


class WallCavityHostInfoV4(object):
    def __init__(self):
        self.kind = "wall_v4"
        self.element = None
        self.element_id = None
        self.level_id = None
        self.level_elevation = 0.0
        self.base_elevation = 0.0
        self.start_point = None
        self.direction = None
        self.normal = None
        self.into_wall = None
        self.length = 0.0
        self.angle = 0.0
        self.target_layer = None
        self.target_depth_from_interior = 0.0
        self.target_layer_label = ""
        self.outer_loop = []
        self.opening_loops = []
        self.openings = []
        self.perimeter_segments = []
        self.wall_solids = []
        self.audit = {}
        self.rejections = []

    def point_at_abs(self, distance_along, z_abs, lateral_offset=0.0):
        from Autodesk.Revit.DB import XYZ

        point = self.start_point + self.direction * distance_along
        if abs(lateral_offset) > 1e-9:
            point = point + self.normal * lateral_offset
        return XYZ(point.X, point.Y, z_abs)

    def point_at(self, distance_along, height, lateral_offset=0.0):
        return self.point_at_abs(
            distance_along,
            self.base_elevation + height,
            lateral_offset,
        )


class WallCavityFramingV4Engine(BaseFramingEngine):
    """Calculate and place wall framing from actual wall face intervals."""

    def calculate_members(self, wall):
        host = self._analyze_wall(wall)
        if host is None:
            return [], None

        occupied = set()
        members = []
        members.extend(self._wall_shape_members(host, occupied))
        members.extend(self._opening_members(host, occupied))
        members.extend(self._infill_members(host, occupied))
        host.audit["member_count"] = len(members)
        return members, host

    # ------------------------------------------------------------------
    # Analysis
    # ------------------------------------------------------------------

    def _analyze_wall(self, wall):
        from Autodesk.Revit.DB import BuiltInParameter, Line, XYZ

        loc = getattr(wall, "Location", None)
        curve = getattr(loc, "Curve", None)
        if curve is None or not isinstance(curve, Line):
            return None

        p0 = curve.GetEndPoint(0)
        p1 = curve.GetEndPoint(1)
        try:
            direction = (p1 - p0).Normalize()
        except Exception:
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

        face_data = self._read_wall_side_face(wall, direction)
        if face_data is None:
            return None

        face, face_normal, loops3, source_side = face_data
        interior_normal = self._orient_as_interior_normal(face_normal, wall, direction)
        if interior_normal is None:
            return None
        into_wall = interior_normal.Multiply(-1.0)

        first_point = _first_loop_point(loops3)
        if first_point is None:
            return None
        p0_on_face = p0 - interior_normal * ((p0 - first_point).DotProduct(interior_normal))

        loops2 = _loops_to_local(loops3, p0_on_face, direction)
        outer_index = _largest_loop_index(loops2)
        if outer_index is None:
            return None

        raw_outer = loops2[outer_index]
        raw_start_d = min(d for d, _ in raw_outer)
        raw_end_d = max(d for d, _ in raw_outer)
        start_d, end_d, domain_source = _trim_domain_to_location_curve(
            raw_start_d,
            raw_end_d,
            curve.Length,
        )
        trimmed_loops2 = _clip_loops_to_d_domain(loops2, start_d, end_d)
        outer_index = _largest_loop_index(trimmed_loops2)
        if outer_index is None:
            return None

        outer = trimmed_loops2[outer_index]
        start_d = min(d for d, _ in outer)
        end_d = max(d for d, _ in outer)
        length = end_d - start_d
        if length < MIN_MEMBER_LENGTH:
            return None

        compound = _get_compound_structure(wall)
        target_depth, target_label, target_layer, total_width = self._target_depth_from_interior(
            compound
        )
        target_offset = target_depth
        if source_side == "exterior":
            target_offset = -max(0.0, total_width - target_depth)
        target_origin = p0_on_face - interior_normal * target_offset
        target_origin = XYZ(target_origin.X, target_origin.Y, base_z)

        wall_solids = _collect_wall_solids(wall)
        if not wall_solids:
            return None

        shifted_loops2 = []
        for loop in trimmed_loops2:
            shifted_loops2.append([(d - start_d, z) for d, z in loop])

        outer_loop = shifted_loops2[outer_index]
        face_opening_pairs = []
        face_opening_candidates = []
        rejected_face_openings = 0
        for index, loop in enumerate(shifted_loops2):
            if index == outer_index:
                continue
            opening = _opening_from_loop(loop, length, base_z, "face")
            if (opening is not None
                    and _face_opening_loop_is_rectangular(loop)
                    and _face_opening_is_actual_void(
                    wall_solids,
                    p0_on_face,
                    direction,
                    interior_normal,
                    total_width,
                    opening,
                    start_d,
                    base_z)):
                face_opening_candidates.append(opening)
                face_opening_pairs.append((opening, loop))
            else:
                rejected_face_openings += 1

        target_loop3 = _target_points_from_local(outer, target_origin, direction)
        perimeter_segments = self._perimeter_segments(
            target_loop3,
            outer_loop,
            start_d,
        )

        hosted_openings = self._hosted_openings(
            wall,
            p0_on_face,
            direction,
            start_d,
            end_d,
            base_z,
        )
        face_openings = [item[0] for item in face_opening_pairs]
        opening_loops = [item[1] for item in face_opening_pairs]
        unmatched_face_openings = len(face_opening_candidates) - len(face_openings)
        openings = _merge_openings(hosted_openings, face_openings, length)

        host = WallCavityHostInfoV4()
        host.element = wall
        host.element_id = wall.Id
        host.level_id = wall.LevelId
        host.level_elevation = level.Elevation
        host.base_elevation = base_z
        host.start_point = target_origin + direction * start_d
        host.direction = direction
        host.normal = interior_normal
        host.into_wall = into_wall
        host.length = length
        host.angle = math.atan2(direction.Y, direction.X)
        host.target_layer = target_layer
        host.target_depth_from_interior = target_depth
        host.target_layer_label = target_label
        host.outer_loop = outer_loop
        host.opening_loops = opening_loops
        host.openings = openings
        host.perimeter_segments = perimeter_segments
        host.wall_solids = wall_solids
        host.audit = {
            "wall_id": _element_id_text(wall.Id),
            "location_length": curve.Length,
            "raw_face_length": raw_end_d - raw_start_d,
            "face_length": length,
            "face_loop_count": len(trimmed_loops2),
            "face_opening_candidate_count": len(face_opening_candidates),
            "face_opening_count": len(face_openings),
            "face_opening_rejected_count": rejected_face_openings,
            "face_opening_unmatched_count": unmatched_face_openings,
            "hosted_opening_count": len(hosted_openings),
            "merged_opening_count": len(openings),
            "perimeter_segment_count": len(perimeter_segments),
            "target_depth": target_depth,
            "target_offset_from_face": target_offset,
            "target_layer_label": target_label,
            "target_layer_index": getattr(target_layer, "index", None),
            "target_layer_width": getattr(target_layer, "width", None),
            "source_side": source_side,
            "domain_source": domain_source,
            "domain_start": start_d,
            "domain_end": end_d,
            "wall_solid_count": len(wall_solids),
            "candidate_count": 0,
            "validated_count": 0,
            "rejected_count": 0,
            "member_count": 0,
        }
        return host

    @staticmethod
    def _read_wall_side_face(wall, direction):
        from Autodesk.Revit.DB import HostObjectUtils, Options, ShellLayerType, Solid

        for shell_layer, side_name in (
                (ShellLayerType.Interior, "interior"),
                (ShellLayerType.Exterior, "exterior")):
            candidates = []
            try:
                refs = HostObjectUtils.GetSideFaces(wall, shell_layer)
            except Exception:
                refs = []
            for reference in refs:
                try:
                    face = wall.GetGeometryObjectFromReference(reference)
                except Exception:
                    face = None
                if face is None:
                    continue
                normal = _face_normal(face)
                if normal is None or abs(normal.Z) > 0.2:
                    continue
                loops = _face_loops_3d(face)
                if not loops:
                    continue
                local = _loops_to_local(loops, _first_loop_point(loops), direction)
                if not local:
                    continue
                area = abs(_polygon_area(local[0]))
                candidates.append((area, face, normal, loops, side_name))
            if candidates:
                candidates.sort(key=lambda item: item[0], reverse=True)
                return candidates[0][1], candidates[0][2], candidates[0][3], candidates[0][4]

        try:
            opts = Options()
            opts.ComputeReferences = True
            geom = wall.get_Geometry(opts)
        except Exception:
            geom = None
        if geom is None:
            return None

        fallback = []
        for geom_obj in geom:
            solids = []
            if isinstance(geom_obj, Solid) and geom_obj.Volume > 0:
                solids.append(geom_obj)
            elif hasattr(geom_obj, "GetInstanceGeometry"):
                try:
                    inst_geom = geom_obj.GetInstanceGeometry()
                except Exception:
                    inst_geom = None
                if inst_geom:
                    for inst_obj in inst_geom:
                        if isinstance(inst_obj, Solid) and inst_obj.Volume > 0:
                            solids.append(inst_obj)
            for solid in solids:
                for face in solid.Faces:
                    normal = _face_normal(face)
                    if normal is None or abs(normal.Z) > 0.2:
                        continue
                    loops = _face_loops_3d(face)
                    if not loops:
                        continue
                    local = _loops_to_local(loops, _first_loop_point(loops), direction)
                    if not local:
                        continue
                    area = abs(_polygon_area(local[0]))
                    fallback.append((area, face, normal, loops, _face_side_from_normal(wall, normal)))

        if not fallback:
            return None
        fallback.sort(key=lambda item: item[0], reverse=True)
        return fallback[0][1], fallback[0][2], fallback[0][3], fallback[0][4]

    @staticmethod
    def _orient_as_interior_normal(face_normal, wall, direction):
        normal = _horizontal_unit(face_normal)
        if normal is None:
            normal = safe_wall_normal(wall, direction)
            if normal is None:
                return None

        exterior = _horizontal_unit(getattr(wall, "Orientation", None))
        if exterior is not None:
            try:
                if normal.DotProduct(exterior) > 0.0:
                    normal = normal.Multiply(-1.0)
            except Exception:
                pass
        return normal

    def _target_depth_from_interior(self, compound):
        layers = _build_compound_layers(self.doc, compound)
        target_layer = _select_target_layer(layers, LAYER_MODE_STRUCTURAL)
        target_layer = _preferred_wall_target_layer(
            layers,
            LAYER_MODE_STRUCTURAL,
            target_layer,
        )
        if compound is None:
            return 0.0, "interior face", target_layer, 0.0

        try:
            raw_layers = list(compound.GetLayers())
        except Exception:
            raw_layers = []
        if not raw_layers:
            return 0.0, "interior face", target_layer, 0.0

        widths = []
        total = 0.0
        for layer in raw_layers:
            width = max(0.0, float(getattr(layer, "Width", 0.0) or 0.0))
            widths.append(width)
            total += width
        if total <= 1e-9:
            return 0.0, "interior face", target_layer, 0.0

        starts = []
        cursor = 0.0
        for width in widths:
            starts.append(cursor)
            cursor += width

        try:
            structural_index = int(compound.StructuralMaterialIndex)
        except Exception:
            structural_index = -1
        if 0 <= structural_index < len(widths) and widths[structural_index] > 1e-9:
            center_from_exterior = starts[structural_index] + widths[structural_index] * 0.5
            return max(0.0, total - center_from_exterior), "structural layer {0}".format(structural_index), target_layer, total

        try:
            first_core = int(compound.GetFirstCoreLayerIndex())
            last_core = int(compound.GetLastCoreLayerIndex())
        except Exception:
            first_core = -1
            last_core = -1
        if 0 <= first_core <= last_core < len(widths):
            core_start = starts[first_core]
            core_end = starts[last_core] + widths[last_core]
            center_from_exterior = (core_start + core_end) * 0.5
            return max(0.0, total - center_from_exterior), "core center", target_layer, total

        thickest = 0
        for index in range(len(widths)):
            if widths[index] > widths[thickest]:
                thickest = index
        center_from_exterior = starts[thickest] + widths[thickest] * 0.5
        return max(0.0, total - center_from_exterior), "thickest layer {0}".format(thickest), target_layer, total

    def _perimeter_segments(self, target_loop3, outer_loop, start_d):
        segments = []
        if len(target_loop3) < 2 or len(outer_loop) < 2:
            return segments

        count = min(len(target_loop3), len(outer_loop))
        for index in range(count):
            p0 = target_loop3[index]
            p1 = target_loop3[(index + 1) % count]
            d0, z0 = outer_loop[index]
            d1, z1 = outer_loop[(index + 1) % count]
            if p0.DistanceTo(p1) < MIN_MEMBER_LENGTH:
                continue

            kind = _classify_perimeter_edge(outer_loop, d0, z0, d1, z1)
            segments.append(WallCavitySegmentV4(kind, p0, p1, d0, d1, z0, z1))
        return segments

    def _hosted_openings(self, wall, origin, direction, start_d, end_d, base_z):
        openings = []
        try:
            insert_ids = wall.FindInserts(True, False, True, False)
        except Exception:
            insert_ids = []
        for insert_id in insert_ids:
            try:
                elem = self.doc.GetElement(insert_id)
            except Exception:
                elem = None
            opening = self._opening_from_insert(
                wall,
                elem,
                origin,
                direction,
                start_d,
                end_d,
                base_z,
            )
            if opening is not None:
                openings.append(opening)
        return openings

    def _opening_from_insert(self, wall, elem, origin, direction, start_d, end_d, base_z):
        if elem is None:
            return None

        from Autodesk.Revit.DB import BuiltInCategory, BuiltInParameter, Opening, Wall, WallKind

        try:
            if isinstance(elem, Opening):
                if not _insert_host_matches_wall(elem, wall):
                    return None
                boundary = elem.BoundaryRect
                if boundary and len(boundary) >= 2:
                    d0 = (boundary[0] - origin).DotProduct(direction)
                    d1 = (boundary[1] - origin).DotProduct(direction)
                    return _opening_from_abs_span(
                        min(d0, d1),
                        max(d0, d1),
                        min(boundary[0].Z, boundary[1].Z),
                        max(boundary[0].Z, boundary[1].Z),
                        start_d,
                        end_d,
                        base_z,
                        "insert",
                    )
        except Exception:
            pass

        if isinstance(elem, Wall):
            try:
                if elem.WallType.Kind != WallKind.Curtain:
                    return None
            except Exception:
                return None
            bbox_opening = _opening_from_element_bbox(
                elem,
                origin,
                direction,
                start_d,
                end_d,
                base_z,
                "curtain",
            )
            if bbox_opening is not None:
                return bbox_opening
            loc = getattr(elem, "Location", None)
            curve = getattr(loc, "Curve", None)
            if curve is None:
                return None
            try:
                d0 = (curve.GetEndPoint(0) - origin).DotProduct(direction)
                d1 = (curve.GetEndPoint(1) - origin).DotProduct(direction)
            except Exception:
                return None
            sill_abs = base_z
            try:
                level = self.doc.GetElement(elem.LevelId)
                if level is not None:
                    sill_abs = level.Elevation
                offset = elem.get_Parameter(BuiltInParameter.WALL_BASE_OFFSET)
                if offset is not None and offset.HasValue:
                    sill_abs += offset.AsDouble()
            except Exception:
                pass
            height = _wall_unconnected_height(elem, inches_to_feet(96.0))
            return _opening_from_abs_span(
                min(d0, d1),
                max(d0, d1),
                sill_abs,
                sill_abs + height,
                start_d,
                end_d,
                base_z,
                "curtain",
            )

        category = getattr(elem, "Category", None)
        category_id = None
        if category is not None:
            category_id = getattr(category.Id, "IntegerValue", getattr(category.Id, "Value", None))
        is_window = category_id == int(BuiltInCategory.OST_Windows)
        is_door = category_id == int(BuiltInCategory.OST_Doors)
        if not is_window and not is_door:
            return None
        if not _insert_host_matches_wall(elem, wall):
            return None

        point = getattr(getattr(elem, "Location", None), "Point", None)
        if point is None:
            return None
        center_d = (point - origin).DotProduct(direction)
        width = _get_opening_width(elem)
        height = _get_opening_height(elem)
        sill = _get_sill_height(elem) if is_window else 0.0
        return _opening_from_abs_span(
            center_d - width * 0.5,
            center_d + width * 0.5,
            base_z + sill,
            base_z + sill + height,
            start_d,
            end_d,
            base_z,
            "family",
        )

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
        bottom_count = int(getattr(self.config, "bottom_plate_count", 1))
        for segment in _segments_by_kind(host, "bottom"):
            gaps = _door_gaps_for_segment(host, segment)
            for plate_index in range(bottom_count):
                z_delta = PLATE_THICKNESS * (plate_index + 0.5)
                for start_d, end_d in _subtract_intervals([(segment.d_min(), segment.d_max())], gaps):
                    start = _point_at_segment_d(segment, start_d, z_delta)
                    end = _point_at_segment_d(segment, end_d, z_delta)
                    members.append(
                        self._member_from_points(
                            host,
                            "BOTTOM_PLATE",
                            start,
                            end,
                            family,
                            type_name,
                            False,
                            PLATE_ROTATION,
                            PLATE_THICKNESS,
                            self._wall_member_depth(host, family, type_name, False),
                        )
                    )
        return [member for member in members if member is not None]

    def _top_plates(self, host):
        members = []
        family = self.config.top_plate_family_name or self.config.bottom_plate_family_name or self.config.stud_family_name
        type_name = self.config.top_plate_type_name or self.config.bottom_plate_type_name or self.config.stud_type_name
        top_count = int(getattr(self.config, "top_plate_count", 2))
        for segment in _segments_by_kind(host, "top"):
            for plate_index in range(top_count):
                offset = (top_count - plate_index - 0.5) * PLATE_THICKNESS
                start = _offset_segment_vertical(segment.p0, -offset)
                end = _offset_segment_vertical(segment.p1, -offset)
                members.append(
                    self._member_from_points(
                        host,
                        "TOP_PLATE",
                        start,
                        end,
                        family,
                        type_name,
                        False,
                        PLATE_ROTATION,
                        PLATE_THICKNESS,
                        self._wall_member_depth(host, family, type_name, False),
                    )
                )
        return [member for member in members if member is not None]

    def _side_studs(self, host, occupied):
        members = []
        for segment in _segments_by_kind(host, "side"):
            d = _side_stud_position(host, (segment.d0 + segment.d1) * 0.5)
            low_z = min(segment.z0, segment.z1) + self._stud_bottom()
            high_z = max(segment.z0, segment.z1) - self._top_plate_stack()
            member = self._vertical_member_at_d(
                host,
                "SIDE_STUD",
                d,
                low_z,
                high_z,
                False,
            )
            if member is not None:
                members.append(member)
                occupied.add(round(d, 4))
        return members

    def _opening_members(self, host, occupied):
        members = []
        header_family = self.config.header_family_name or self.config.stud_family_name
        header_type = self.config.header_type_name or self.config.stud_type_name
        header_depth = self._header_depth(header_family, header_type)
        header_width = self._type_width(header_family, header_type, False)
        plate_family = self.config.bottom_plate_family_name or self.config.stud_family_name
        plate_type = self.config.bottom_plate_type_name or self.config.stud_type_name
        plate_depth = self._wall_member_depth(host, plate_family, plate_type, False)

        for opening in host.openings:
            left = opening.left_edge
            right = opening.right_edge
            if right - left < MIN_MEMBER_LENGTH:
                continue

            if getattr(self.config, "include_king_studs", True):
                for d in (left - STUD_THICKNESS * 1.5, right + STUD_THICKNESS * 1.5):
                    if d <= 0.0 or d >= host.length or _near(d, occupied, OPENING_STUD_COLLISION_TOL):
                        continue
                    member = self._vertical_member_at_d(host, "KING_STUD", d, None, None, False)
                    if member is not None:
                        members.append(member)
                        occupied.add(round(d, 4))

            if getattr(self.config, "include_jack_studs", True):
                for d in (left - STUD_THICKNESS * 0.5, right + STUD_THICKNESS * 0.5):
                    if d <= 0.0 or d >= host.length or _near(d, occupied, OPENING_STUD_COLLISION_TOL):
                        continue
                    member = self._vertical_member_at_d(
                        host,
                        "JACK_STUD",
                        d,
                        None,
                        host.base_elevation + opening.head_height,
                        False,
                    )
                    if member is not None:
                        members.append(member)
                        occupied.add(round(d, 4))

            span_left = max(0.0, left - STUD_THICKNESS)
            span_right = min(host.length, right + STUD_THICKNESS)
            header_stack_top = host.base_elevation + opening.head_height
            if span_right - span_left >= MIN_MEMBER_LENGTH:
                bottom_plate_center = (
                    host.base_elevation
                    + opening.head_height
                    + PLATE_THICKNESS * 0.5
                )
                start = host.point_at_abs(span_left, bottom_plate_center)
                end = host.point_at_abs(span_right, bottom_plate_center)
                members.append(
                    self._member_from_points(
                        host,
                        "HEADER_BOTTOM_PLATE",
                        start,
                        end,
                        plate_family,
                        plate_type,
                        False,
                        PLATE_ROTATION,
                        PLATE_THICKNESS,
                        plate_depth,
                    )
                )

                header_bottom = host.base_elevation + opening.head_height + PLATE_THICKNESS
                header_center = header_bottom + header_depth * 0.5
                for lateral in _header_ply_lateral_offsets(
                        host,
                        int(getattr(self.config, "header_count", MIN_HEADER_PLY_COUNT)),
                        header_width):
                    start = host.point_at_abs(span_left, header_center, lateral)
                    end = host.point_at_abs(span_right, header_center, lateral)
                    members.append(
                        self._member_from_points(
                            host,
                            "HEADER",
                            start,
                            end,
                            header_family,
                            header_type,
                            False,
                            HEADER_ROTATION,
                            header_depth,
                            header_width,
                        )
                    )
                header_top = header_bottom + header_depth
                top_plate_center = header_top + PLATE_THICKNESS * 0.5
                start = host.point_at_abs(span_left, top_plate_center)
                end = host.point_at_abs(span_right, top_plate_center)
                members.append(
                    self._member_from_points(
                        host,
                        "HEADER_TOP_PLATE",
                        start,
                        end,
                        plate_family,
                        plate_type,
                        False,
                        PLATE_ROTATION,
                        PLATE_THICKNESS,
                        plate_depth,
                    )
                )
                header_stack_top = header_top + PLATE_THICKNESS

            if getattr(self.config, "include_cripple_studs", True):
                top_bound = _top_bound_at_d(host.outer_loop, (left + right) * 0.5)
                if top_bound is not None:
                    top_z = top_bound - self._top_plate_stack()
                    if header_stack_top < top_z - MIN_MEMBER_LENGTH:
                        members.extend(
                            self._cripples(host, left, right, header_stack_top, top_z, occupied)
                        )

            if opening.is_window and opening.sill_height > self._stud_bottom():
                sill_center = host.base_elevation + opening.sill_height - PLATE_THICKNESS * 0.5
                start = host.point_at_abs(left, sill_center)
                end = host.point_at_abs(right, sill_center)
                members.append(
                    self._member_from_points(
                        host,
                        "SILL_PLATE",
                        start,
                        end,
                        self.config.bottom_plate_family_name or self.config.stud_family_name,
                        self.config.bottom_plate_type_name or self.config.stud_type_name,
                        False,
                        PLATE_ROTATION,
                        PLATE_THICKNESS,
                        self._wall_member_depth(
                            host,
                            self.config.bottom_plate_family_name or self.config.stud_family_name,
                            self.config.bottom_plate_type_name or self.config.stud_type_name,
                            False,
                        ),
                    )
                )
                if getattr(self.config, "include_cripple_studs", True):
                    bottom_z = _bottom_bound_at_d(host.outer_loop, (left + right) * 0.5)
                    if bottom_z is None:
                        bottom_z = host.base_elevation
                    bottom_z += self._stud_bottom()
                    top_z = host.base_elevation + opening.sill_height - PLATE_THICKNESS
                    if top_z > bottom_z + MIN_MEMBER_LENGTH:
                        members.extend(
                            self._cripples(host, left, right, bottom_z, top_z, occupied)
                        )

        return [member for member in members if member is not None]

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

        family = self.config.bottom_plate_family_name or self.config.stud_family_name
        type_name = self.config.bottom_plate_type_name or self.config.stud_type_name
        z_abs = host.base_elevation + self._stud_bottom() + interval
        max_top = max(z for _, z in host.outer_loop) - self._top_plate_stack()
        while z_abs < max_top - PLATE_THICKNESS:
            intervals = _horizontal_intervals(host.outer_loop, host.opening_loops, z_abs)
            intervals = _subtract_intervals(
                intervals,
                _opening_gaps_at_z(host.openings, z_abs - host.base_elevation),
            )
            for start_d, end_d in intervals:
                if end_d - start_d < MIN_MEMBER_LENGTH:
                    continue
                start = host.point_at_abs(start_d, z_abs + PLATE_THICKNESS * 0.5)
                end = host.point_at_abs(end_d, z_abs + PLATE_THICKNESS * 0.5)
                members.append(
                    self._member_from_points(
                        host,
                        "MID_PLATE",
                        start,
                        end,
                        family,
                        type_name,
                        False,
                        PLATE_ROTATION,
                        PLATE_THICKNESS,
                        self._wall_member_depth(host, family, type_name, False),
                    )
                )
            z_abs += interval
        return [member for member in members if member is not None]

    def _regular_studs(self, host, occupied):
        members = []
        spacing = self.config.stud_spacing_ft
        if spacing <= 0.0:
            return members

        for segment in _segments_by_kind(host, "bottom"):
            d_min = max(0.0, segment.d_min())
            d_max = min(host.length, segment.d_max())
            d = _next_spacing_station(d_min, spacing)
            while d < d_max - STUD_THICKNESS * 0.5:
                if not _in_opening_zone(host.openings, d) and not _near(d, occupied, STUD_THICKNESS):
                    intervals = _vertical_intervals(host.outer_loop, [], d)
                    for bottom_z, top_z in intervals:
                        bottom_z += self._stud_bottom()
                        top_z -= self._top_plate_stack()
                        member = self._vertical_member_at_d(
                            host,
                            "STUD",
                            d,
                            bottom_z,
                            top_z,
                            False,
                        )
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
        family = self.config.bottom_plate_family_name or self.config.stud_family_name
        type_name = self.config.bottom_plate_type_name or self.config.stud_type_name
        sorted_positions = sorted(occupied)
        for index in range(len(sorted_positions) - 1):
            left = sorted_positions[index]
            right = sorted_positions[index + 1]
            if right - left < MIN_MEMBER_LENGTH or right - left > spacing * 1.5:
                continue
            mid = (left + right) * 0.5
            bounds = _vertical_bounds(host.outer_loop, mid)
            if bounds is None:
                continue
            bottom_z, top_z = bounds
            clear_height = top_z - bottom_z - self._stud_bottom() - self._top_plate_stack()
            if clear_height <= 8.0:
                continue
            z_abs = bottom_z + self._stud_bottom() + clear_height * 0.5
            start = host.point_at_abs(left + STUD_THICKNESS * 0.5, z_abs)
            end = host.point_at_abs(right - STUD_THICKNESS * 0.5, z_abs)
            members.append(
                self._member_from_points(
                    host,
                    "BLOCKING",
                    start,
                    end,
                    family,
                    type_name,
                    False,
                    PLATE_ROTATION,
                    PLATE_THICKNESS,
                    self._wall_member_depth(host, family, type_name, False),
                )
            )
        return [member for member in members if member is not None]

    # ------------------------------------------------------------------
    # Member factories and validation
    # ------------------------------------------------------------------

    def _vertical_member_at_d(self, host, kind, d, bottom_abs_z, top_abs_z,
                              respect_openings):
        intervals = _vertical_intervals(
            host.outer_loop,
            host.opening_loops if respect_openings else [],
            d,
        )
        if not intervals:
            return None

        target_bottom = bottom_abs_z
        target_top = top_abs_z
        best = None
        for low, high in intervals:
            framed_low = low + self._stud_bottom()
            framed_high = high - self._top_plate_stack()
            if target_bottom is not None:
                framed_low = max(framed_low, target_bottom)
            if target_top is not None:
                framed_high = min(framed_high, target_top)
            if framed_high - framed_low < MIN_MEMBER_LENGTH:
                continue
            candidate_height = framed_high - framed_low
            if best is None or candidate_height > best[1] - best[0]:
                best = (framed_low, framed_high)

        if best is None:
            return None
        start = host.point_at_abs(d, best[0])
        end = host.point_at_abs(d, best[1])
        return self._member_from_points(
            host,
            kind,
            start,
            end,
            self.config.stud_family_name,
            self.config.stud_type_name,
            True,
            host.angle,
            STUD_THICKNESS,
            self._wall_member_depth(
                host,
                self.config.stud_family_name,
                self.config.stud_type_name,
                True,
            ),
        )

    def _member_from_points(self, host, member_type, start, end, family, type_name,
                            is_column, rotation, section_height, section_depth):
        if start is None or end is None:
            return None
        try:
            if start.DistanceTo(end) < MIN_MEMBER_LENGTH:
                return None
        except Exception:
            return None

        member = FramingMember(member_type, start, end)
        member.member_type = member_type
        member.family_name = family
        member.type_name = type_name
        member.rotation = rotation
        member.is_column = bool(is_column)
        member.host_kind = host.kind
        member.host_id = host.element_id
        if host.target_layer is not None:
            member.layer_index = host.target_layer.index
        member.disallow_end_joins = not bool(is_column)
        member.section_height = max(STUD_THICKNESS, float(section_height or STUD_THICKNESS))
        member.section_depth = max(STUD_THICKNESS, float(section_depth or DEFAULT_LUMBER_DEPTH))

        return self._validated_member(host, member)

    def _validated_member(self, host, member):
        host.audit["candidate_count"] = host.audit.get("candidate_count", 0) + 1
        valid, reason = self._member_inside_wall_solid(host, member)
        if valid:
            host.audit["validated_count"] = host.audit.get("validated_count", 0) + 1
            return member
        host.audit["rejected_count"] = host.audit.get("rejected_count", 0) + 1
        if not host.audit.get("first_rejection"):
            host.audit["first_rejection"] = reason
        if len(host.rejections) < 12:
            host.rejections.append((member.member_type, reason))
        return None

    def _member_inside_wall_solid(self, host, member):
        if not host.wall_solids:
            return False, "no wall solid"
        if getattr(member, "member_type", None) in OPENING_FRAME_MEMBER_TYPES:
            return self._opening_frame_member_inside_wall_solid(host, member)
        lines = self._validation_sample_lines(host, member)
        if not lines:
            return False, "no sample lines"
        for start, end in lines:
            if not _line_inside_any_solid(host.wall_solids, start, end):
                return False, "sample outside wall solid"
        return True, None

    def _opening_frame_member_inside_wall_solid(self, host, member):
        start = member.start_point
        end = member.end_point
        if start is None or end is None:
            return False, "missing member endpoints"
        if not _line_inside_any_solid(host.wall_solids, start, end):
            return False, "opening frame centerline outside wall solid"

        edge_lines = self._validation_sample_lines(host, member)
        if not edge_lines:
            return True, None
        failed = 0
        for edge_start, edge_end in edge_lines:
            if not _line_inside_any_solid(host.wall_solids, edge_start, edge_end):
                failed += 1
        if failed > max(4, len(edge_lines) - 3):
            return False, "opening frame mostly outside wall solid"
        return True, None

    def _validation_sample_lines(self, host, member):
        start = member.start_point
        end = member.end_point
        if start is None or end is None:
            return []

        if getattr(member, "is_column", False):
            width_axis = host.direction
            depth_axis = host.into_wall
            half_width = max(STUD_THICKNESS * 0.5, member.section_height * 0.5)
            half_depth = max(STUD_THICKNESS * 0.5, member.section_depth * 0.5)
        else:
            try:
                beam_axis = (end - start).Normalize()
            except Exception:
                return []
            depth_axis = host.into_wall
            width_axis = _normalize(beam_axis.CrossProduct(depth_axis))
            if width_axis is None:
                from Autodesk.Revit.DB import XYZ
                width_axis = XYZ.BasisZ
            half_width = max(STUD_THICKNESS * 0.5, member.section_height * 0.5)
            half_depth = max(STUD_THICKNESS * 0.5, member.section_depth * 0.5)

        scale = 0.96
        offsets = [
            (0.0, 0.0),
            (half_width * scale, 0.0),
            (-half_width * scale, 0.0),
            (0.0, half_depth * scale),
            (0.0, -half_depth * scale),
            (half_width * scale, half_depth * scale),
            (half_width * scale, -half_depth * scale),
            (-half_width * scale, half_depth * scale),
            (-half_width * scale, -half_depth * scale),
        ]
        lines = []
        for width_offset, depth_offset in offsets:
            offset_vec = width_axis * width_offset + depth_axis * depth_offset
            lines.append((start + offset_vec, end + offset_vec))
        return lines

    def _wall_member_depth(self, host, family, type_name, is_column):
        target_layer = getattr(host, "target_layer", None)
        layer_width = float(getattr(target_layer, "width", 0.0) or 0.0)
        symbol_depth = self._type_depth(family, type_name, is_column)
        if layer_width > STUD_THICKNESS:
            return min(max(symbol_depth, STUD_THICKNESS), layer_width)
        return max(symbol_depth, STUD_THICKNESS)

    def _type_depth(self, family, type_name, is_column):
        from Autodesk.Revit.DB import BuiltInCategory

        category = (
            BuiltInCategory.OST_StructuralColumns
            if is_column
            else BuiltInCategory.OST_StructuralFraming
        )
        symbol = None
        if family and type_name:
            symbol = find_family_symbol(self.doc, family, type_name, category)
        if symbol is None:
            return DEFAULT_LUMBER_DEPTH
        for name in ("d", "Depth", "Nominal Depth", "Height"):
            value = _lookup_double(symbol, name)
            if value is not None and value > 0.0:
                return value
        return DEFAULT_LUMBER_DEPTH

    def _header_depth(self, family, type_name):
        depth = self._type_depth(family, type_name, False)
        if depth > 0.0:
            return depth
        return DEFAULT_HEADER_DEPTH

    def _type_width(self, family, type_name, is_column):
        from Autodesk.Revit.DB import BuiltInCategory

        category = (
            BuiltInCategory.OST_StructuralColumns
            if is_column
            else BuiltInCategory.OST_StructuralFraming
        )
        symbol = None
        if family and type_name:
            symbol = find_family_symbol(self.doc, family, type_name, category)
        if symbol is None:
            return DEFAULT_LUMBER_WIDTH
        for name in ("b", "Width", "Nominal Width"):
            value = _lookup_double(symbol, name)
            if value is not None and value > 0.0:
                return value
        return DEFAULT_LUMBER_WIDTH

    def _stud_bottom(self):
        return PLATE_THICKNESS * int(getattr(self.config, "bottom_plate_count", 1))

    def _top_plate_stack(self):
        return PLATE_THICKNESS * int(getattr(self.config, "top_plate_count", 2))

    def _cripples(self, host, left, right, bottom_abs_z, top_abs_z, occupied):
        members = []
        spacing = self.config.stud_spacing_ft
        if spacing <= 0.0 or top_abs_z - bottom_abs_z < MIN_MEMBER_LENGTH:
            return members
        d = _next_spacing_station(left, spacing)
        while d < right - STUD_THICKNESS * 0.5:
            if left + STUD_THICKNESS * 0.5 < d < right - STUD_THICKNESS * 0.5:
                if not _near(d, occupied, STUD_THICKNESS):
                    member = self._vertical_member_at_d(
                        host,
                        "CRIPPLE_STUD",
                        d,
                        bottom_abs_z,
                        top_abs_z,
                        False,
                    )
                    if member is not None:
                        members.append(member)
            d += spacing
        return members


# ----------------------------------------------------------------------
# Geometry helpers
# ----------------------------------------------------------------------


def _collect_wall_solids(wall):
    from Autodesk.Revit.DB import GeometryInstance, Options, Solid, ViewDetailLevel

    try:
        opts = Options()
        opts.ComputeReferences = True
        opts.DetailLevel = ViewDetailLevel.Fine
        geom = wall.get_Geometry(opts)
    except Exception:
        geom = None
    if geom is None:
        return []

    solids = []
    for geom_obj in geom:
        if isinstance(geom_obj, Solid) and geom_obj.Volume > 1e-9:
            solids.append(geom_obj)
            continue
        if isinstance(geom_obj, GeometryInstance):
            try:
                inst_geom = geom_obj.GetInstanceGeometry()
            except Exception:
                inst_geom = None
            if inst_geom is None:
                continue
            for sub in inst_geom:
                if isinstance(sub, Solid) and sub.Volume > 1e-9:
                    solids.append(sub)
    return solids


def _face_loops_3d(face):
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
                tessellated = list(curve.Tessellate())
            except Exception:
                tessellated = []
            for point in tessellated:
                if points and _same_xyz(points[-1], point):
                    continue
                points.append(point)
        if len(points) >= 2 and _same_xyz(points[0], points[-1]):
            points = points[:-1]
        if len(points) >= 3:
            loops.append(points)
    return loops


def _face_normal(face):
    try:
        bbox = face.GetBoundingBox()
        uv = None
        if bbox is not None:
            from Autodesk.Revit.DB import UV
            uv = UV(
                (bbox.Min.U + bbox.Max.U) * 0.5,
                (bbox.Min.V + bbox.Max.V) * 0.5,
            )
        if uv is not None:
            normal = face.ComputeNormal(uv)
        else:
            normal = face.FaceNormal
        return normal.Normalize()
    except Exception:
        return None


def _loops_to_local(loops3, origin, direction):
    result = []
    if origin is None:
        return result
    for loop in loops3:
        local = []
        for point in loop:
            d = (point - origin).DotProduct(direction)
            candidate = (d, point.Z)
            if local:
                prev = local[-1]
                if abs(prev[0] - candidate[0]) < 1e-7 and abs(prev[1] - candidate[1]) < 1e-7:
                    continue
            local.append(candidate)
        if len(local) >= 3:
            result.append(local)
    return result


def _trim_domain_to_location_curve(face_start, face_end, location_length):
    if location_length is None or location_length < MIN_MEMBER_LENGTH:
        return face_start, face_end, "face"

    trim_start = max(face_start, 0.0)
    trim_end = min(face_end, float(location_length))
    if trim_end - trim_start >= MIN_MEMBER_LENGTH:
        return trim_start, trim_end, "location"
    return face_start, face_end, "face"


def _clip_loops_to_d_domain(loops, start_d, end_d):
    clipped = []
    for loop in loops:
        next_loop = _clip_loop_to_min_d(loop, start_d)
        next_loop = _clip_loop_to_max_d(next_loop, end_d)
        next_loop = _clean_loop_points(next_loop)
        if len(next_loop) >= 3 and abs(_polygon_area(next_loop)) > MIN_MEMBER_LENGTH * MIN_MEMBER_LENGTH:
            clipped.append(next_loop)
    return clipped


def _clip_loop_to_min_d(loop, min_d):
    return _clip_loop_by_d(loop, min_d, True)


def _clip_loop_to_max_d(loop, max_d):
    return _clip_loop_by_d(loop, max_d, False)


def _clip_loop_by_d(loop, limit, keep_greater):
    if not loop:
        return []

    def inside(point):
        if keep_greater:
            return point[0] >= limit - GEOM_TOL
        return point[0] <= limit + GEOM_TOL

    result = []
    previous = loop[-1]
    previous_inside = inside(previous)
    for current in loop:
        current_inside = inside(current)
        if current_inside != previous_inside:
            result.append(_interpolate_loop_point_at_d(previous, current, limit))
        if current_inside:
            result.append(current)
        previous = current
        previous_inside = current_inside
    return result


def _interpolate_loop_point_at_d(first, second, d_value):
    d0, z0 = first
    d1, z1 = second
    denom = d1 - d0
    if abs(denom) < 1e-9:
        return (d_value, z0)
    t = (d_value - d0) / denom
    t = max(0.0, min(1.0, t))
    return (d_value, z0 + (z1 - z0) * t)


def _clean_loop_points(loop):
    result = []
    for d, z in loop:
        point = (d, z)
        if result:
            prev = result[-1]
            if abs(prev[0] - point[0]) < 1e-7 and abs(prev[1] - point[1]) < 1e-7:
                continue
        result.append(point)
    if len(result) > 1:
        first = result[0]
        last = result[-1]
        if abs(first[0] - last[0]) < 1e-7 and abs(first[1] - last[1]) < 1e-7:
            result = result[:-1]
    return result


def _target_points_from_local(loop, target_origin, direction):
    result = []
    for d, z in loop:
        result.append(_point_from_local(target_origin, direction, d, z))
    return result


def _point_from_local(origin, direction, d, z_abs):
    from Autodesk.Revit.DB import XYZ

    point = origin + direction * d
    return XYZ(point.X, point.Y, z_abs)


def _face_opening_is_actual_void(wall_solids, face_origin, direction,
                                 interior_normal, total_width, opening,
                                 domain_start, base_z):
    if not wall_solids:
        return False

    probe_depth = max(float(total_width or 0.0), DEFAULT_LUMBER_DEPTH)
    probe_depth += inches_to_feet(2.0)
    max_void_solid = max(inches_to_feet(0.75), probe_depth * 0.20)
    sample_count = 0
    void_count = 0
    for d_factor in (0.25, 0.5, 0.75):
        sample_d = domain_start + opening.left_edge
        sample_d += (opening.right_edge - opening.left_edge) * d_factor
        for z_factor in (0.25, 0.5, 0.75):
            sample_z = base_z + opening.sill_height
            sample_z += (opening.head_height - opening.sill_height) * z_factor
            center = _point_from_local(face_origin, direction, sample_d, sample_z)
            start = center - interior_normal * probe_depth
            end = center + interior_normal * probe_depth
            inside = _solid_intersection_length(wall_solids, start, end)
            sample_count += 1
            if inside <= max_void_solid:
                void_count += 1
    if sample_count <= 0:
        return False
    return void_count >= sample_count - 1


def _face_opening_loop_is_rectangular(loop):
    if not loop:
        return False
    left = min(d for d, _ in loop)
    right = max(d for d, _ in loop)
    bottom = min(z for _, z in loop)
    top = max(z for _, z in loop)
    if right - left < inches_to_feet(6.0):
        return False
    if top - bottom < inches_to_feet(6.0):
        return False

    edge_tol = inches_to_feet(1.0)
    corners = {
        "lb": False,
        "rb": False,
        "rt": False,
        "lt": False,
    }
    off_edge = 0
    for d, z in loop:
        on_left = abs(d - left) <= edge_tol
        on_right = abs(d - right) <= edge_tol
        on_bottom = abs(z - bottom) <= edge_tol
        on_top = abs(z - top) <= edge_tol
        if not (on_left or on_right or on_bottom or on_top):
            off_edge += 1
        if on_left and on_bottom:
            corners["lb"] = True
        if on_right and on_bottom:
            corners["rb"] = True
        if on_right and on_top:
            corners["rt"] = True
        if on_left and on_top:
            corners["lt"] = True
    return off_edge == 0 and all(corners.values())


def _insert_host_matches_wall(elem, wall):
    host = getattr(elem, "Host", None)
    if host is None:
        return True
    return _same_element_id(getattr(host, "Id", None), getattr(wall, "Id", None))


def _same_element_id(first, second):
    return _element_id_text(first) == _element_id_text(second)


def _header_ply_lateral_offsets(host, header_count, ply_width=None):
    count = max(MIN_HEADER_PLY_COUNT, int(header_count or MIN_HEADER_PLY_COUNT))
    if count <= 1:
        return [0.0]

    ply_width = max(DEFAULT_LUMBER_WIDTH, float(ply_width or DEFAULT_LUMBER_WIDTH))
    target_layer = getattr(host, "target_layer", None)
    layer_width = float(getattr(target_layer, "width", 0.0) or 0.0)
    if layer_width <= ply_width:
        layer_width = max(DEFAULT_LUMBER_DEPTH, ply_width * count)

    default_step = ply_width + HEADER_PLY_SPACER
    max_step = max(ply_width, (layer_width - ply_width) / float(count - 1))
    step = min(default_step, max_step)
    center = (count - 1) * 0.5
    return [(index - center) * step for index in range(count)]


def _classify_perimeter_edge(outer_loop, d0, z0, d1, z1):
    vertical = abs(d1 - d0) <= inches_to_feet(0.5)
    height = abs(z1 - z0)
    if vertical and height > inches_to_feet(6.0):
        return "side"

    mid_d = (d0 + d1) * 0.5
    mid_z = (z0 + z1) * 0.5
    bounds = _vertical_bounds(outer_loop, mid_d)
    if bounds is not None and mid_z <= bounds[0] + inches_to_feet(2.0):
        return "bottom"
    return "top"


def _opening_from_loop(loop, length, base_z, source):
    if not loop:
        return None
    left = max(0.0, min(d for d, _ in loop))
    right = min(length, max(d for d, _ in loop))
    sill_abs = min(z for _, z in loop)
    head_abs = max(z for _, z in loop)
    if right - left < inches_to_feet(6.0):
        return None
    if head_abs - sill_abs < inches_to_feet(6.0):
        return None
    if source == "face" and (left <= PROFILE_EDGE_TOL or right >= length - PROFILE_EDGE_TOL):
        return None
    return WallCavityOpeningV4(
        left,
        right,
        sill_abs - base_z,
        head_abs - base_z,
        sill_abs > base_z + PLATE_THICKNESS * 2.0,
        source,
    )


def _opening_from_abs_span(left_abs, right_abs, sill_abs, head_abs,
                           start_d, end_d, base_z, source):
    length = max(0.0, end_d - start_d)
    left = max(0.0, left_abs - start_d)
    right = min(length, right_abs - start_d)
    if right - left < inches_to_feet(6.0):
        return None
    if head_abs - sill_abs < inches_to_feet(6.0):
        return None
    return WallCavityOpeningV4(
        left,
        right,
        sill_abs - base_z,
        head_abs - base_z,
        sill_abs > base_z + PLATE_THICKNESS * 2.0,
        source,
    )


def _opening_from_element_bbox(element, origin, direction, start_d, end_d,
                               base_z, source):
    try:
        bbox = element.get_BoundingBox(None)
    except Exception:
        bbox = None
    if bbox is None:
        return None

    corners = _bbox_corners(bbox)
    if not corners:
        return None
    distances = []
    zs = []
    for point in corners:
        try:
            distances.append((point - origin).DotProduct(direction))
            zs.append(point.Z)
        except Exception:
            pass
    if not distances or not zs:
        return None

    return _opening_from_abs_span(
        min(distances),
        max(distances),
        min(zs),
        max(zs),
        start_d,
        end_d,
        base_z,
        source,
    )


def _bbox_corners(bbox):
    try:
        from Autodesk.Revit.DB import XYZ
        min_pt = bbox.Min
        max_pt = bbox.Max
        return [
            XYZ(min_pt.X, min_pt.Y, min_pt.Z),
            XYZ(min_pt.X, min_pt.Y, max_pt.Z),
            XYZ(min_pt.X, max_pt.Y, min_pt.Z),
            XYZ(min_pt.X, max_pt.Y, max_pt.Z),
            XYZ(max_pt.X, min_pt.Y, min_pt.Z),
            XYZ(max_pt.X, min_pt.Y, max_pt.Z),
            XYZ(max_pt.X, max_pt.Y, min_pt.Z),
            XYZ(max_pt.X, max_pt.Y, max_pt.Z),
        ]
    except Exception:
        return []


def _merge_openings(hosted, face, length):
    merged = []
    for source in (hosted or [], face or []):
        for opening in source:
            left = max(0.0, opening.left_edge)
            right = min(length, opening.right_edge)
            if right - left < inches_to_feet(6.0):
                continue
            duplicate = False
            for existing in merged:
                if _same_or_overlapping_opening(existing, left, right):
                    duplicate = True
                    existing.left_edge = min(existing.left_edge, left)
                    existing.right_edge = max(existing.right_edge, right)
                    existing.sill_height = min(existing.sill_height, opening.sill_height)
                    existing.head_height = max(existing.head_height, opening.head_height)
                    existing.is_window = existing.is_window and opening.is_window
                    existing.is_door = not existing.is_window
                    break
            if not duplicate:
                merged.append(
                    WallCavityOpeningV4(
                        left,
                        right,
                        opening.sill_height,
                        opening.head_height,
                        opening.is_window,
                        opening.source,
                    )
                )
    merged.sort(key=lambda item: item.left_edge)
    return merged


def _same_or_overlapping_opening(existing, left, right):
    if (abs(existing.left_edge - left) < STUD_THICKNESS and
            abs(existing.right_edge - right) < STUD_THICKNESS):
        return True
    overlap = min(existing.right_edge, right) - max(existing.left_edge, left)
    if overlap <= 0.0:
        return False
    existing_width = existing.right_edge - existing.left_edge
    candidate_width = right - left
    smaller = min(existing_width, candidate_width)
    if smaller <= 0.0:
        return False
    return overlap / smaller >= 0.65


def _segments_by_kind(host, kind):
    return [segment for segment in host.perimeter_segments if segment.kind == kind]


def _door_gaps_for_segment(host, segment):
    gaps = []
    d0 = segment.d_min()
    d1 = segment.d_max()
    for opening in host.openings:
        if opening.is_window:
            continue
        left = max(d0, opening.left_edge)
        right = min(d1, opening.right_edge)
        if right - left > MIN_MEMBER_LENGTH:
            gaps.append((left, right))
    return gaps


def _opening_gaps_at_z(openings, z_height):
    gaps = []
    for opening in openings:
        if opening.sill_height - GEOM_TOL <= z_height <= opening.head_height + GEOM_TOL:
            gaps.append((opening.left_edge, opening.right_edge))
    return gaps


def _side_stud_position(host, d):
    if d <= STUD_THICKNESS:
        return min(host.length * 0.5, STUD_THICKNESS * 0.5)
    if d >= host.length - STUD_THICKNESS:
        return max(0.0, host.length - STUD_THICKNESS * 0.5)
    return d


def _horizontal_intervals(outer_loop, opening_loops, z_abs):
    intervals = _scan_loop_intervals(outer_loop, z_abs, "horizontal")
    for opening_loop in opening_loops:
        intervals = _subtract_intervals(
            intervals,
            _scan_loop_intervals(opening_loop, z_abs, "horizontal"),
        )
    return intervals


def _vertical_intervals(outer_loop, opening_loops, d):
    intervals = _scan_loop_intervals(outer_loop, d, "vertical")
    for opening_loop in opening_loops:
        intervals = _subtract_intervals(
            intervals,
            _scan_loop_intervals(opening_loop, d, "vertical"),
        )
    return intervals


def _vertical_bounds(outer_loop, d):
    intervals = _vertical_intervals(outer_loop, [], d)
    if not intervals:
        return None
    return intervals[0][0], intervals[-1][1]


def _bottom_bound_at_d(outer_loop, d):
    bounds = _vertical_bounds(outer_loop, d)
    if bounds is None:
        return None
    return bounds[0]


def _top_bound_at_d(outer_loop, d):
    bounds = _vertical_bounds(outer_loop, d)
    if bounds is None:
        return None
    return bounds[1]


def _scan_loop_intervals(loop, coord, mode):
    values = []
    count = len(loop)
    if count < 3:
        return []
    for index in range(count):
        d0, z0 = loop[index]
        d1, z1 = loop[(index + 1) % count]
        if mode == "vertical":
            value = _edge_intersection(d0, z0, d1, z1, coord, "d")
        else:
            value = _edge_intersection(d0, z0, d1, z1, coord, "z")
        if value is not None:
            values.append(value)

    values = _unique_sorted(values)
    intervals = []
    index = 0
    while index + 1 < len(values):
        start = values[index]
        end = values[index + 1]
        if end - start > MIN_MEMBER_LENGTH:
            intervals.append((start, end))
        index += 2
    return intervals


def _edge_intersection(d0, z0, d1, z1, coord, axis):
    if axis == "d":
        a0, a1 = d0, d1
        b0, b1 = z0, z1
    else:
        a0, a1 = z0, z1
        b0, b1 = d0, d1

    if abs(a1 - a0) < 1e-9:
        return None
    low = min(a0, a1)
    high = max(a0, a1)
    if coord < low - 1e-9 or coord >= high - 1e-9:
        return None
    t = (coord - a0) / (a1 - a0)
    if t < -1e-9 or t > 1.0 + 1e-9:
        return None
    return b0 + (b1 - b0) * t


def _subtract_intervals(intervals, gaps):
    result = list(intervals)
    for gap_start, gap_end in gaps:
        next_result = []
        for start, end in result:
            if gap_end <= start or gap_start >= end:
                next_result.append((start, end))
                continue
            if gap_start > start:
                next_result.append((start, gap_start))
            if gap_end < end:
                next_result.append((gap_end, end))
        result = next_result
    return [(start, end) for start, end in result if end - start >= MIN_MEMBER_LENGTH]


def _line_inside_any_solid(solids, start, end):
    try:
        total = start.DistanceTo(end)
    except Exception:
        return False
    if total < MIN_MEMBER_LENGTH:
        return False
    inside = _solid_intersection_length(solids, start, end)
    return inside >= total - VALIDATION_TOL


def _solid_intersection_length(solids, start, end):
    from Autodesk.Revit.DB import Line, SolidCurveIntersectionMode, SolidCurveIntersectionOptions

    try:
        line = Line.CreateBound(start, end)
    except Exception:
        return 0.0
    inside = 0.0
    for solid in solids:
        try:
            options = SolidCurveIntersectionOptions()
            try:
                options.ResultType = SolidCurveIntersectionMode.CurveSegmentsInside
            except Exception:
                pass
            result = solid.IntersectWithCurve(line, options)
            count = result.SegmentCount
        except Exception:
            continue
        for index in range(count):
            try:
                segment = result.GetCurveSegment(index)
                inside += segment.Length
            except Exception:
                continue
    return inside


def _point_at_segment_d(segment, d, z_delta):
    denom = segment.d1 - segment.d0
    if abs(denom) < 1e-9:
        t = 0.0
    else:
        t = (d - segment.d0) / denom
    t = max(0.0, min(1.0, t))
    from Autodesk.Revit.DB import XYZ
    return XYZ(
        segment.p0.X + (segment.p1.X - segment.p0.X) * t,
        segment.p0.Y + (segment.p1.Y - segment.p0.Y) * t,
        segment.p0.Z + (segment.p1.Z - segment.p0.Z) * t + z_delta,
    )


def _offset_segment_vertical(point, delta):
    from Autodesk.Revit.DB import XYZ
    return XYZ(point.X, point.Y, point.Z + delta)


def _next_spacing_station(start_d, spacing):
    if spacing <= 0.0:
        return start_d
    return math.ceil((start_d + STUD_THICKNESS * 0.5) / spacing) * spacing


def _in_opening_zone(openings, d):
    for opening in openings:
        if opening.left_edge - STUD_THICKNESS * 3.0 <= d <= opening.right_edge + STUD_THICKNESS * 3.0:
            return True
    return False


def _near(value, occupied, tolerance):
    for existing in occupied:
        if abs(value - existing) < tolerance:
            return True
    return False


def _polygon_area(points):
    area = 0.0
    for index in range(len(points)):
        x0, y0 = points[index]
        x1, y1 = points[(index + 1) % len(points)]
        area += (x0 * y1) - (x1 * y0)
    return area * 0.5


def _largest_loop_index(loops):
    best_index = None
    best_area = 0.0
    for index, loop in enumerate(loops):
        area = abs(_polygon_area(loop))
        if best_index is None or area > best_area:
            best_index = index
            best_area = area
    return best_index


def _first_loop_point(loops):
    for loop in loops:
        if loop:
            return loop[0]
    return None


def _same_xyz(first, second):
    try:
        return first.DistanceTo(second) < 1e-7
    except Exception:
        return False


def _face_side_from_normal(wall, normal):
    exterior = _horizontal_unit(getattr(wall, "Orientation", None))
    face_normal = _horizontal_unit(normal)
    if exterior is None or face_normal is None:
        return "unknown"
    try:
        if face_normal.DotProduct(exterior) > 0.0:
            return "exterior"
    except Exception:
        return "unknown"
    return "interior"


def _horizontal_unit(vector):
    if vector is None:
        return None
    try:
        length = math.sqrt(vector.X * vector.X + vector.Y * vector.Y)
        if length < 1e-9:
            return None
        from Autodesk.Revit.DB import XYZ
        return XYZ(vector.X / length, vector.Y / length, 0.0)
    except Exception:
        return None


def _normalize(vector):
    if vector is None:
        return None
    try:
        length = vector.GetLength()
        if length < 1e-9:
            return None
        return vector.Normalize()
    except Exception:
        return None


def _unique_sorted(values):
    result = []
    for value in sorted(values):
        if not result or abs(value - result[-1]) > 1e-7:
            result.append(value)
    return result


def _lookup_double(element, name):
    try:
        param = element.LookupParameter(name)
    except Exception:
        param = None
    if param is None or not param.HasValue:
        return None
    try:
        return param.AsDouble()
    except Exception:
        return None


def _wall_unconnected_height(wall, default):
    from Autodesk.Revit.DB import BuiltInParameter

    try:
        param = wall.get_Parameter(BuiltInParameter.WALL_USER_HEIGHT_PARAM)
        if param is not None and param.HasValue:
            value = param.AsDouble()
            if value > 0.0:
                return value
    except Exception:
        pass
    return default


def _element_id_text(element_id):
    if element_id is None:
        return None
    value = getattr(element_id, "IntegerValue", getattr(element_id, "Value", None))
    if value is None:
        return None
    return str(value)
