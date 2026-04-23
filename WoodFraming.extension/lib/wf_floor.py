# -*- coding: utf-8 -*-
"""Floor framing engine using the shared host-local framing core."""

from wf_geometry import FramingMember, inches_to_feet
from wf_host import analyze_floor_host
from wf_placement import BaseFramingEngine


MIN_MEMBER_LENGTH = inches_to_feet(1.0)


class FloorFramingEngine(BaseFramingEngine):
    """Calculates and places joists and rim joists for a floor."""

    def calculate_members(self, floor):
        """Calculate framing members for a floor."""
        floor_info = analyze_floor_host(self.doc, floor, self.config)
        if floor_info is None:
            return [], None

        members = []
        members.extend(self._calc_joists(floor_info))
        members.extend(self._calc_rim_joists(floor_info))
        return members, floor_info

    def _calc_joists(self, floor_info):
        """Place joists along the shorter span, clipped to the host profile."""
        members = []
        spacing = self.config.stud_spacing_ft
        if spacing <= 0.0:
            return members

        min_x, max_x, min_y, max_y = floor_info.bounds
        span_x = max_x - min_x
        span_y = max_y - min_y
        if span_x < MIN_MEMBER_LENGTH or span_y < MIN_MEMBER_LENGTH:
            return members

        if span_x <= span_y:
            coords = self._interior_coords(min_y, max_y, spacing)
            for coord in coords:
                intervals = floor_info.scanline_intervals("y", coord)
                for start_x, end_x in intervals:
                    if end_x - start_x < MIN_MEMBER_LENGTH:
                        continue
                    start_pt = floor_info.point_at(start_x, coord)
                    end_pt = floor_info.point_at(end_x, coord)
                    member = FramingMember(FramingMember.STUD, start_pt, end_pt)
                    member.member_type = "JOIST"
                    member.family_name = self.config.stud_family_name
                    member.type_name = self.config.stud_type_name
                    self._apply_member_rule(member, floor_info)
                    members.append(member)
        else:
            coords = self._interior_coords(min_x, max_x, spacing)
            for coord in coords:
                intervals = floor_info.scanline_intervals("x", coord)
                for start_y, end_y in intervals:
                    if end_y - start_y < MIN_MEMBER_LENGTH:
                        continue
                    start_pt = floor_info.point_at(coord, start_y)
                    end_pt = floor_info.point_at(coord, end_y)
                    member = FramingMember(FramingMember.STUD, start_pt, end_pt)
                    member.member_type = "JOIST"
                    member.family_name = self.config.stud_family_name
                    member.type_name = self.config.stud_type_name
                    self._apply_member_rule(member, floor_info)
                    members.append(member)

        return members

    def _calc_rim_joists(self, floor_info):
        """Place rim joists along each boundary segment."""
        members = []

        for loop in floor_info.boundary_loops_local:
            count = len(loop)
            for index in range(count):
                start_local = loop[index]
                end_local = loop[(index + 1) % count]
                dx = end_local[0] - start_local[0]
                dy = end_local[1] - start_local[1]
                if (dx * dx + dy * dy) ** 0.5 < MIN_MEMBER_LENGTH:
                    continue

                start_pt = floor_info.point_at(start_local[0], start_local[1])
                end_pt = floor_info.point_at(end_local[0], end_local[1])
                member = FramingMember(FramingMember.BOTTOM_PLATE, start_pt, end_pt)
                member.member_type = "RIM_JOIST"
                member.family_name = (
                    self.config.bottom_plate_family_name or self.config.stud_family_name
                )
                member.type_name = (
                    self.config.bottom_plate_type_name or self.config.stud_type_name
                )
                self._apply_member_rule(member, floor_info)
                members.append(member)

        return members

    def _apply_member_rule(self, member, floor_info):
        """Attach floor host placement metadata to a generated member.

        Joists and rim joists use the default upright orientation, so
        section_normal is left as None to skip cross-section rotation.
        """
        member.host_kind = floor_info.kind
        member.host_id = floor_info.element_id
        if floor_info.target_layer is not None:
            member.layer_index = floor_info.target_layer.index

    @staticmethod
    def _interior_coords(min_value, max_value, spacing):
        """Generate interior scan coordinates, with a center fallback."""
        coords = []
        coord = min_value + spacing
        while coord < max_value - 1e-9:
            coords.append(coord)
            coord += spacing
        if not coords:
            coords.append((min_value + max_value) / 2.0)
        return coords

