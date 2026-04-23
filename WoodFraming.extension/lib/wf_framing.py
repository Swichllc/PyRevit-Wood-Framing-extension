# -*- coding: utf-8 -*-
"""Wall framing engine.

Calculates stud, plate, header, sill, king-stud, jack-stud and cripple
positions for a straight Revit wall and delegates placement to the shared
BaseFramingEngine.

Rotation values (STRUCTURAL_BEND_DIR_ANGLE) follow Revit conventions.
"""

import math

from wf_geometry import FramingMember, inches_to_feet
from wf_host import analyze_wall_host
from wf_placement import BaseFramingEngine
from wf_wall_joins import build_wall_join_plan


MIN_MEMBER_LENGTH = inches_to_feet(1.0)
STUD_OVERLAP_TOL = inches_to_feet(1.5)
STUD_THICKNESS = inches_to_feet(1.5)  # 1.5" typical
MID_PLATE_INTERVAL = 8.0  # feet

def _plate_rotation():
    """Flat plate rotation is 0.0"""
    return 0.0

def _header_rotation(wall_info):
    """Headers span horizontally and are placed on-edge (-pi/2)."""
    return -math.pi / 2.0


class WallFramingEngine(BaseFramingEngine):
    """Calculates wood framing members for a Revit wall."""

    def __init__(self, doc, config):
        BaseFramingEngine.__init__(self, doc, config)
        self._plate_thickness = inches_to_feet(1.5)
        self._stud_thickness = inches_to_feet(1.5)

    def calculate_members(self, wall):
        wall_info = analyze_wall_host(self.doc, wall, self.config)
        if wall_info is None:
            return [], None

        openings = wall_info.openings
        try:
            join_plan = build_wall_join_plan(
                self.doc,
                wall_info,
                self.config,
                self._stud_thickness,
            )
        except Exception:
            join_plan = None
        occupied = set()
        members = []
        members.extend(self._calc_bottom_plates(wall_info, openings))
        members.extend(self._calc_mid_plates(wall_info, openings))
        members.extend(self._calc_top_plates(wall_info, openings))
        members.extend(self._calc_join_studs(wall_info, openings, occupied, join_plan))
        members.extend(self._calc_opening_framing(wall_info, openings, occupied))
        members.extend(self._calc_regular_studs(wall_info, openings, occupied))
        return members, wall_info

    def _calc_bottom_plates(self, wall_info, openings):
        members = []
        segments = self._plate_segments(wall_info, openings, split_at_doors=True)
        half_t = self._plate_thickness / 2.0

        for i in range(self.config.bottom_plate_count):
            plate_z = self._plate_thickness * i + half_t
            for seg_start, seg_end in segments:
                if seg_end - seg_start < MIN_MEMBER_LENGTH:
                    continue
                start_pt = self._point_on_wall(wall_info, seg_start, plate_z)
                end_pt = self._point_on_wall(wall_info, seg_end, plate_z)
                m = FramingMember(FramingMember.BOTTOM_PLATE, start_pt, end_pt)
                m.family_name = self.config.bottom_plate_family_name or self.config.stud_family_name
                m.type_name = self.config.bottom_plate_type_name or self.config.stud_type_name
                self._tag(m, wall_info, _plate_rotation())
                members.append(m)

        return members

    def _calc_top_plates(self, wall_info, openings):
        members = []
        half_t = self._plate_thickness / 2.0

        start_connected, end_connected = self._connected_wall_ends(wall_info)
        trim_start = self._plate_thickness if start_connected else 0.0
        trim_end = self._plate_thickness if end_connected else 0.0
        seg_start = trim_start
        seg_end = wall_info.length - trim_end
        if seg_end - seg_start < MIN_MEMBER_LENGTH:
            return members

        for i in range(self.config.top_plate_count):
            if wall_info.is_sloped_top:
                offset = self._plate_thickness * i + half_t
                start_h = wall_info.height_at(seg_start)
                end_h = wall_info.height_at(seg_end)
                start_z = start_h - self._plate_thickness * self.config.top_plate_count + offset
                end_z = end_h - self._plate_thickness * self.config.top_plate_count + offset
            else:
                plate_z = self._stud_top_z(wall_info) + self._plate_thickness * i + half_t
                start_z = plate_z
                end_z = plate_z

            start_pt = self._point_on_wall(wall_info, seg_start, start_z)
            end_pt = self._point_on_wall(wall_info, seg_end, end_z)
            if end_pt.DistanceTo(start_pt) < MIN_MEMBER_LENGTH:
                continue

            m = FramingMember(FramingMember.TOP_PLATE, start_pt, end_pt)
            m.family_name = self.config.top_plate_family_name or self.config.stud_family_name
            m.type_name = self.config.top_plate_type_name or self.config.stud_type_name
            self._tag(m, wall_info, _plate_rotation())
            members.append(m)

        return members

    def _calc_mid_plates(self, wall_info, openings):
        """Place optional horizontal mid plates from the stud baseline upward."""
        members = []
        if not bool(getattr(self.config, "include_mid_plates", True)):
            return members

        z_values = self._mid_plate_z_values(wall_info)
        if not z_values:
            return members

        half_t = self._plate_thickness / 2.0
        segments = self._plate_segments(
            wall_info,
            openings,
            split_at_doors=False,
            split_at_openings=True,
        )

        family_name = self.config.bottom_plate_family_name or self.config.stud_family_name
        type_name = self.config.bottom_plate_type_name or self.config.stud_type_name

        for plate_z in z_values:
            center_z = plate_z + half_t
            for seg_start, seg_end in segments:
                if seg_end - seg_start < MIN_MEMBER_LENGTH:
                    continue
                start_pt = self._point_on_wall(wall_info, seg_start, center_z)
                end_pt = self._point_on_wall(wall_info, seg_end, center_z)
                member = FramingMember(FramingMember.BOTTOM_PLATE, start_pt, end_pt)
                member.member_type = "MID_PLATE"
                member.family_name = family_name
                member.type_name = type_name
                self._tag(member, wall_info, _plate_rotation())
                members.append(member)

        return members

    def _calc_regular_studs(self, wall_info, openings, occupied=None):
        members = []
        spacing = self.config.stud_spacing_ft
        if spacing <= 0.0:
            return members

        stud_bottom = self._stud_bottom_z()

        if occupied is None:
            occupied = self._opening_occupied_positions(openings)

        positions = []
        dist = spacing
        while dist < wall_info.length - self._stud_thickness * 0.5:
            positions.append(dist)
            dist += spacing

        for pos in positions:
            if self._is_within_opening(pos, openings):
                continue
            if self._near_occupied(pos, occupied):
                continue
            stud_top = self._stud_top_z(wall_info, pos)
            if stud_top - stud_bottom < MIN_MEMBER_LENGTH:
                continue
            members.append(
                self._make_stud_member(
                    FramingMember.STUD,
                    wall_info,
                    pos,
                    stud_bottom,
                    stud_top,
                )
            )
            occupied.add(round(pos, 4))

        return members

    def _calc_join_studs(self, wall_info, openings, occupied, join_plan=None):
        members = []
        stud_bottom = self._stud_bottom_z()

        phys_start, phys_end = self._physical_end_distances(
            wall_info, stud_bottom + self._stud_thickness)
        if phys_end - phys_start < MIN_MEMBER_LENGTH:
            phys_start, phys_end = (0.0, wall_info.length)

        run_length = max(0.0, phys_end - phys_start)
        edge_backset = self._stud_thickness * 0.5 + self._wrapped_layer_backset(wall_info)
        max_backset = max(0.0, (run_length - self._stud_thickness) * 0.5)
        edge_backset = min(edge_backset, max_backset)

        def _line_to_physical(line_dist):
            return self._line_to_physical_distance(
                line_dist,
                wall_info.length,
                phys_start,
                phys_end,
            )

        def _place_stud_piece(candidate):
            cand = max(phys_start, min(phys_end, candidate))

            if self._near_occupied(cand, occupied):
                return True

            if self._is_within_opening(cand, openings):
                return False

            stud_top = self._stud_top_z(wall_info, cand)
            if stud_top - stud_bottom < MIN_MEMBER_LENGTH:
                return False

            members.append(
                self._make_stud_member(
                    FramingMember.STUD,
                    wall_info,
                    cand,
                    stud_bottom,
                    stud_top,
                )
            )
            occupied.add(round(cand, 4))
            return True

        def _try_place_end(target_pos, end_index, max_steps=7):
            step = self._stud_thickness if end_index == 0 else -self._stud_thickness
            for i in range(0, max_steps):
                cand = target_pos + step * i
                if _place_stud_piece(cand):
                    return True
            return False

        def _try_place_intersection(target_pos, max_steps=2):
            if self._is_within_opening(target_pos, openings):
                return False

            offsets = [0.0]
            for i in range(1, max_steps + 1):
                delta = self._stud_thickness * i
                offsets.append(delta)
                offsets.append(-delta)
            for offset in offsets:
                if _place_stud_piece(target_pos + offset):
                    return True
            return False

        for end_index in (0, 1):
            has_revit_join = self._has_revit_join_at_end(wall_info, end_index)

            if not has_revit_join:
                free_target = (
                    phys_start + edge_backset
                    if end_index == 0
                    else phys_end - edge_backset
                )
                _try_place_end(free_target, end_index)
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
                    phys_start + edge_backset
                    if end_index == 0
                    else phys_end - edge_backset
                )
                _try_place_end(free_target, end_index)
                continue

            for line_pos in line_positions:
                if line_pos < -1e-9 or line_pos > wall_info.length + 1e-9:
                    continue
                _try_place_end(_line_to_physical(line_pos), end_index)

        if join_plan is not None:
            intersections = getattr(join_plan, "intersections", []) or []
            for intersection in intersections:
                raw_positions = list(getattr(intersection, "positions", []) or [])
                if not raw_positions:
                    distance = getattr(intersection, "distance", None)
                    if distance is None:
                        continue
                    raw_positions = [
                        distance - (self._stud_thickness * 0.5),
                        distance + (self._stud_thickness * 0.5),
                    ]

                for line_pos in raw_positions:
                    if line_pos <= self._stud_thickness:
                        continue
                    if line_pos >= wall_info.length - self._stud_thickness:
                        continue
                    _try_place_intersection(_line_to_physical(line_pos))

        return members

    def _calc_opening_framing(self, wall_info, openings, occupied):
        members = []
        stud_bottom = self._stud_bottom_z()
        header_depth = self._get_header_depth()

        for op in openings:
            stud_top = self._stud_top_z(wall_info, op.distance_along_wall)
            
            left_jack_c = op.left_edge - self._stud_thickness * 0.5
            right_jack_c = op.right_edge + self._stud_thickness * 0.5
            left_king_c = op.left_edge - self._stud_thickness * 1.5
            right_king_c = op.right_edge + self._stud_thickness * 1.5
            header_bottom_z = op.head_height

            # Kings
            if self.config.include_king_studs:
                for c in (left_king_c, right_king_c):
                    if 0 <= c <= wall_info.length:
                        if self._near_occupied(c, occupied):
                            continue
                        s_top = self._stud_top_z(wall_info, c)
                        s = self._point_on_wall(wall_info, c, stud_bottom)
                        e = self._point_on_wall(wall_info, c, s_top)
                        m = FramingMember(FramingMember.KING_STUD, s, e)
                        m.family_name = self.config.stud_family_name
                        m.type_name = self.config.stud_type_name
                        self._tag(m, wall_info, wall_info.angle, is_column=True)
                        members.append(m)
                        occupied.add(round(c, 4))

            # Jacks
            if self.config.include_jack_studs and header_bottom_z > stud_bottom:
                for c in (left_jack_c, right_jack_c):
                    if 0 <= c <= wall_info.length:
                        if self._near_occupied(c, occupied):
                            continue
                        s = self._point_on_wall(wall_info, c, stud_bottom)
                        e = self._point_on_wall(wall_info, c, header_bottom_z)
                        m = FramingMember(FramingMember.JACK_STUD, s, e)
                        m.family_name = self.config.stud_family_name
                        m.type_name = self.config.stud_type_name
                        self._tag(m, wall_info, wall_info.angle, is_column=True)
                        members.append(m)
                        occupied.add(round(c, 4))

            # Headers
            header_count = getattr(self.config, "header_count", 2)
            span_start = op.left_edge - self._stud_thickness
            span_end = op.right_edge + self._stud_thickness
            
            if (span_end - span_start) >= MIN_MEMBER_LENGTH:
                header_center_z = header_bottom_z + header_depth / 2.0
                b = self._plate_thickness
                
                for h_idx in range(header_count):
                    lateral = (h_idx - (header_count - 1) / 2.0) * b
                    s = self._point_on_wall(wall_info, span_start, header_center_z, lateral)
                    e = self._point_on_wall(wall_info, span_end, header_center_z, lateral)
                    m = FramingMember(FramingMember.HEADER, s, e)
                    m.family_name = self.config.header_family_name or self.config.stud_family_name
                    m.type_name = self.config.header_type_name or self.config.stud_type_name
                    self._tag(m, wall_info, _header_rotation(wall_info))
                    members.append(m)

            # Cripples above
            header_top_z = header_bottom_z + header_depth
            if self.config.include_cripple_studs and header_top_z < stud_top:
                members.extend(
                    self._calc_cripples(
                        wall_info,
                        op.left_edge,
                        op.right_edge,
                        header_top_z,
                        stud_top,
                        occupied,
                    )
                )

            # Sill & Cripples below
            if op.is_window and op.sill_height > stud_bottom:
                sill_z = op.sill_height
                sill_center_z = sill_z - self._plate_thickness / 2.0
                if (op.right_edge - op.left_edge) >= MIN_MEMBER_LENGTH:
                    s = self._point_on_wall(wall_info, op.left_edge, sill_center_z)
                    e = self._point_on_wall(wall_info, op.right_edge, sill_center_z)
                    m = FramingMember(FramingMember.SILL_PLATE, s, e)
                    m.family_name = getattr(self.config, "sill_plate_family_name", None) or self.config.stud_family_name
                    m.type_name = getattr(self.config, "sill_plate_type_name", None) or self.config.stud_type_name
                    self._tag(m, wall_info, _plate_rotation())
                    members.append(m)

                if self.config.include_cripple_studs and sill_z - self._plate_thickness > stud_bottom:
                    members.extend(
                        self._calc_cripples(
                            wall_info,
                            op.left_edge,
                            op.right_edge,
                            stud_bottom,
                            sill_z - self._plate_thickness,
                            occupied,
                        )
                    )

        return members

    def _calc_cripples(self, wall_info, left, right, bottom_z, top_z, occupied=None):
        members = []
        spacing = self.config.stud_spacing_ft
        if spacing <= 0.0 or top_z - bottom_z < MIN_MEMBER_LENGTH:
            return members

        dist = left + spacing
        while dist < right - STUD_OVERLAP_TOL:
            if occupied is not None and self._near_occupied(dist, occupied):
                dist += spacing
                continue
            s = self._point_on_wall(wall_info, dist, bottom_z)
            e = self._point_on_wall(wall_info, dist, top_z)
            m = FramingMember(FramingMember.CRIPPLE_STUD, s, e)
            m.family_name = self.config.stud_family_name
            m.type_name = self.config.stud_type_name
            self._tag(m, wall_info, wall_info.angle, is_column=True)
            members.append(m)
            if occupied is not None:
                occupied.add(round(dist, 4))
            dist += spacing
        return members

    def _physical_end_distances(self, wall_info, sample_z):
        """Get physical wall run limits from 3D solid intersections."""
        wall = wall_info.element
        if wall is None:
            return (0.0, wall_info.length)

        try:
            from Autodesk.Revit.DB import (
                GeometryInstance,
                Line,
                Options,
                Solid,
                SolidCurveIntersectionOptions,
                ViewDetailLevel,
            )

            opts = Options()
            opts.ComputeReferences = False
            opts.DetailLevel = ViewDetailLevel.Fine
            geom_elem = wall.get_Geometry(opts)
        except Exception:
            return (0.0, wall_info.length)

        solids = []
        for gobj in geom_elem:
            if isinstance(gobj, Solid) and gobj.Volume > 0:
                solids.append(gobj)
            elif isinstance(gobj, GeometryInstance):
                try:
                    for sub in gobj.GetInstanceGeometry():
                        if isinstance(sub, Solid) and sub.Volume > 0:
                            solids.append(sub)
                except Exception:
                    pass

        if not solids:
            return (0.0, wall_info.length)

        p0 = wall_info.point_at(0.0, 0.0)
        ray_origin = wall_info.point_at(0.0, sample_z)
        ray_pad = max(2.0, wall_info.length * 0.25, self._stud_thickness * 8.0)
        ray_start = ray_origin - wall_info.direction * ray_pad
        ray_end = ray_origin + wall_info.direction * (wall_info.length + ray_pad)
        try:
            ray = Line.CreateBound(ray_start, ray_end)
        except Exception:
            return (0.0, wall_info.length)

        hit_min = None
        hit_max = None
        for solid in solids:
            try:
                result = solid.IntersectWithCurve(
                    ray, SolidCurveIntersectionOptions())
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
                d_a = (p_a - p0).DotProduct(wall_info.direction)
                d_b = (p_b - p0).DotProduct(wall_info.direction)
                lo = min(d_a, d_b)
                hi = max(d_a, d_b)
                if hit_min is None or lo < hit_min:
                    hit_min = lo
                if hit_max is None or hi > hit_max:
                    hit_max = hi

        if hit_min is None or hit_max is None:
            return (0.0, wall_info.length)

        min_d = max(0.0, min(wall_info.length, hit_min))
        max_d = max(0.0, min(wall_info.length, hit_max))
        if max_d - min_d < MIN_MEMBER_LENGTH:
            return (0.0, wall_info.length)
        return (min_d, max_d)

    def _get_header_depth(self):
        family = self.config.header_family_name or self.config.stud_family_name
        tname = self.config.header_type_name or self.config.stud_type_name
        try:
            depth = self.get_type_depth(family, tname)
        except AttributeError:
            depth = None
        if depth is not None:
            return depth
        return self._plate_thickness

    def _mid_plate_z_values(self, wall_info):
        """Return z values (local wall-space) for configured mid-plate tiers."""
        interval = float(getattr(self.config, "mid_plate_interval_ft", MID_PLATE_INTERVAL))
        if interval <= 0.0:
            return []

        stud_bottom = self._stud_bottom_z()
        top_start = self._stud_top_z(wall_info, 0.0)
        top_end = self._stud_top_z(wall_info, wall_info.length)
        min_top = min(top_start, top_end)
        total_stud_height = min_top - stud_bottom

        if total_stud_height <= interval + self._plate_thickness * 1.5:
            return []

        z_values = []
        tier = 1
        while True:
            z_value = stud_bottom + tier * interval
            if z_value < min_top - self._plate_thickness:
                z_values.append(z_value)
                tier += 1
                continue
            break

        return z_values

    def _plate_segments(self, wall_info, openings, split_at_doors=False, split_at_openings=False):
        segments = [(0.0, wall_info.length)]
        if not split_at_doors and not split_at_openings:
            return segments

        for op in openings:
            if split_at_openings:
                gap_left = op.left_edge
                gap_right = op.right_edge
            else:
                if not op.is_door:
                    continue
                gap_left = op.left_edge - self._stud_thickness
                gap_right = op.right_edge + self._stud_thickness

            new_segments = []
            for seg_start, seg_end in segments:
                if gap_right <= seg_start or gap_left >= seg_end:
                    new_segments.append((seg_start, seg_end))
                    continue
                if gap_left > seg_start:
                    new_segments.append((seg_start, gap_left))
                if gap_right < seg_end:
                    new_segments.append((gap_right, seg_end))
            segments = new_segments

        return segments

    def _stud_bottom_z(self):
        return self._plate_thickness * self.config.bottom_plate_count

    def _stud_top_z(self, wall_info, distance_along=None):
        if distance_along is not None:
            h = wall_info.height_at(distance_along)
        else:
            h = wall_info.height
        return h - self._plate_thickness * self.config.top_plate_count

    def _point_on_wall(self, wall_info, dist_along, height, extra_offset=0.0):
        return wall_info.point_at(dist_along, height, extra_offset)

    def _make_stud_member(self, member_type, wall_info, dist_along, bottom_z, top_z):
        start_pt = self._point_on_wall(wall_info, dist_along, bottom_z)
        end_pt = self._point_on_wall(wall_info, dist_along, top_z)
        member = FramingMember(member_type, start_pt, end_pt)
        member.family_name = self.config.stud_family_name
        member.type_name = self.config.stud_type_name
        self._tag(member, wall_info, wall_info.angle, is_column=True)
        return member

    def _opening_occupied_positions(self, openings):
        occupied = set()
        for op in openings:
            if self.config.include_king_studs:
                occupied.add(round(op.left_edge - self._stud_thickness * 1.5, 4))
                occupied.add(round(op.right_edge + self._stud_thickness * 1.5, 4))
            if self.config.include_jack_studs:
                occupied.add(round(op.left_edge - self._stud_thickness * 0.5, 4))
                occupied.add(round(op.right_edge + self._stud_thickness * 0.5, 4))
        return occupied

    def _tag(self, member, wall_info, rotation, is_column=False):
        """Attach host metadata, rotation and structural type to a member."""
        member.rotation = rotation
        member.is_column = is_column
        member.host_kind = wall_info.kind
        member.host_id = wall_info.element_id
        if wall_info.target_layer is not None:
            member.layer_index = wall_info.target_layer.index

    def _connected_wall_ends(self, wall_info):
        """Return endpoint connectivity to neighboring straight walls."""
        from Autodesk.Revit.DB import FilteredElementCollector, Wall, Line

        start_conn = False
        end_conn = False
        tol = self._plate_thickness * 2.0

        try:
            walls = (
                FilteredElementCollector(self.doc)
                .OfClass(Wall)
                .WhereElementIsNotElementType()
            )
        except Exception:
            return (False, False)

        for other in walls:
            try:
                if other.Id == wall_info.element_id:
                    continue
                oloc = other.Location
                if oloc is None:
                    continue
                oc = oloc.Curve
                if not isinstance(oc, Line):
                    continue
                for end_idx in (0, 1):
                    pt = oc.GetEndPoint(end_idx)
                    if self._xy_distance(pt, wall_info.start_point) < tol:
                        start_conn = True
                    if self._xy_distance(pt, wall_info.end_point) < tol:
                        end_conn = True
                if start_conn and end_conn:
                    break
            except Exception:
                continue

        return (start_conn, end_conn)

    @staticmethod
    def _is_within_opening(dist, openings):
        for op in openings:
            if (op.left_edge - STUD_THICKNESS * 3.0) <= dist <= (op.right_edge + STUD_THICKNESS * 3.0):
                return True
        return False

    def _near_occupied(self, dist, occupied_positions):
        rounded = round(dist, 4)
        for occ in occupied_positions:
            if abs(rounded - occ) < self._stud_thickness:
                return True
        return False

    @staticmethod
    def _line_to_physical_distance(line_dist, line_length, phys_start, phys_end):
        if line_length <= 1e-9:
            return (phys_start + phys_end) * 0.5
        ratio = max(0.0, min(1.0, line_dist / line_length))
        return phys_start + (phys_end - phys_start) * ratio

    def _wrapped_layer_backset(self, wall_info):
        base_info = getattr(wall_info, "wall_info", None)
        wall_width = getattr(base_info, "width", 0.0)
        target_layer = getattr(wall_info, "target_layer", None)
        target_width = getattr(target_layer, "width", 0.0)
        if wall_width <= 1e-9 or target_width <= 1e-9:
            return 0.0
        return max(0.0, (wall_width - target_width) * 0.5)

    def _has_revit_join_at_end(self, wall_info, end_index):
        wall = getattr(wall_info, "element", None)
        if wall is None:
            return True
        try:
            from Autodesk.Revit.DB import WallUtils
            return bool(WallUtils.IsWallJoinAllowedAtEnd(wall, end_index))
        except Exception:
            return True

    @staticmethod
    def _xy_distance(a, b):
        dx = a.X - b.X
        dy = a.Y - b.Y
        return math.sqrt(dx * dx + dy * dy)
