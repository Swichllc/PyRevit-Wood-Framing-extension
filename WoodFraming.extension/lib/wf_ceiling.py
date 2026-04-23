# -*- coding: utf-8 -*-
"""Ceiling framing engine using the shared host-local framing core."""

import re

from wf_geometry import FramingMember, inches_to_feet
from wf_config import (
    CEILING_DIRECTION_AUTO,
    CEILING_DIRECTION_X,
    CEILING_DIRECTION_Y,
    CEILING_PLACEMENT_ABOVE,
    CEILING_PLACEMENT_CENTER,
    LUMBER_ACTUAL,
)
from wf_host import analyze_ceiling_host
from wf_placement import BaseFramingEngine


MIN_MEMBER_LENGTH = inches_to_feet(1.0)


class CeilingFramingEngine(BaseFramingEngine):
    """Calculates and places joists and rim joists for a ceiling."""

    def calculate_members(self, ceiling):
        """Calculate framing members for a ceiling."""
        ceiling_info = analyze_ceiling_host(self.doc, ceiling, self.config)
        if ceiling_info is None:
            return [], None

        members = []
        members.extend(self._calc_joists(ceiling_info))
        members.extend(self._calc_rim_joists(ceiling_info))
        return members, ceiling_info

    def _calc_joists(self, ceiling_info):
        """Place joists along the shorter span, clipped to the host profile."""
        members = []
        spacing = self.config.stud_spacing_ft
        if spacing <= 0.0:
            return members

        min_x, max_x, min_y, max_y = ceiling_info.bounds
        span_x = max_x - min_x
        span_y = max_y - min_y
        if span_x < MIN_MEMBER_LENGTH or span_y < MIN_MEMBER_LENGTH:
            return members

        layout_axis = self._resolve_layout_axis(span_x, span_y)

        if layout_axis == "x":
            coords = self._layout_coords(min_y, max_y, spacing)
            for coord in coords:
                intervals = ceiling_info.scanline_intervals("y", coord)
                for start_x, end_x in intervals:
                    if end_x - start_x < MIN_MEMBER_LENGTH:
                        continue
                    start_pt = self._member_point(
                        ceiling_info,
                        start_x,
                        coord,
                        self.config.stud_family_name,
                        self.config.stud_type_name,
                    )
                    end_pt = self._member_point(
                        ceiling_info,
                        end_x,
                        coord,
                        self.config.stud_family_name,
                        self.config.stud_type_name,
                    )
                    member = FramingMember(FramingMember.STUD, start_pt, end_pt)
                    member.member_type = "CEILING_JOIST"
                    member.family_name = self.config.stud_family_name
                    member.type_name = self.config.stud_type_name
                    self._apply_member_rule(member, ceiling_info)
                    members.append(member)
        else:
            coords = self._layout_coords(min_x, max_x, spacing)
            for coord in coords:
                intervals = ceiling_info.scanline_intervals("x", coord)
                for start_y, end_y in intervals:
                    if end_y - start_y < MIN_MEMBER_LENGTH:
                        continue
                    start_pt = self._member_point(
                        ceiling_info,
                        coord,
                        start_y,
                        self.config.stud_family_name,
                        self.config.stud_type_name,
                    )
                    end_pt = self._member_point(
                        ceiling_info,
                        coord,
                        end_y,
                        self.config.stud_family_name,
                        self.config.stud_type_name,
                    )
                    member = FramingMember(FramingMember.STUD, start_pt, end_pt)
                    member.member_type = "CEILING_JOIST"
                    member.family_name = self.config.stud_family_name
                    member.type_name = self.config.stud_type_name
                    self._apply_member_rule(member, ceiling_info)
                    members.append(member)

        return members

    def _calc_rim_joists(self, ceiling_info):
        """Place rim joists along each boundary segment."""
        members = []
        family_name = (
            self.config.bottom_plate_family_name or self.config.stud_family_name
        )
        type_name = (
            self.config.bottom_plate_type_name or self.config.stud_type_name
        )

        for loop in ceiling_info.boundary_loops_local:
            count = len(loop)
            for index in range(count):
                start_local = loop[index]
                end_local = loop[(index + 1) % count]
                dx = end_local[0] - start_local[0]
                dy = end_local[1] - start_local[1]
                if (dx * dx + dy * dy) ** 0.5 < MIN_MEMBER_LENGTH:
                    continue

                start_pt = self._member_point(
                    ceiling_info,
                    start_local[0],
                    start_local[1],
                    family_name,
                    type_name,
                )
                end_pt = self._member_point(
                    ceiling_info,
                    end_local[0],
                    end_local[1],
                    family_name,
                    type_name,
                )
                member = FramingMember(FramingMember.BOTTOM_PLATE, start_pt, end_pt)
                member.member_type = "CEILING_RIM_JOIST"
                member.family_name = family_name
                member.type_name = type_name
                self._apply_member_rule(member, ceiling_info)
                members.append(member)

        return members

    def _apply_member_rule(self, member, ceiling_info):
        """Attach ceiling host placement metadata to a generated member."""
        member.host_kind = ceiling_info.kind
        member.host_id = ceiling_info.element_id
        if ceiling_info.target_layer is not None:
            member.layer_index = ceiling_info.target_layer.index

    @staticmethod
    def _layout_coords(min_value, max_value, spacing):
        """Generate centered framing coordinates from both edges inward."""
        span = max_value - min_value
        if span <= 1e-9 or spacing <= 1e-9:
            return []

        interval_count = int(span / spacing)
        if interval_count <= 1:
            return [(min_value + max_value) / 2.0]

        edge_gap = (span - (interval_count * spacing)) / 2.0
        coords = []
        coord = min_value + edge_gap + spacing
        limit = max_value - edge_gap - 1e-9
        while coord < limit:
            coords.append(coord)
            coord += spacing
        if not coords:
            coords.append((min_value + max_value) / 2.0)
        return coords

    def _resolve_layout_axis(self, span_x, span_y):
        """Return the joist run axis in host-local coordinates."""
        mode = getattr(self.config, "ceiling_direction_mode", CEILING_DIRECTION_AUTO)
        if mode == CEILING_DIRECTION_X:
            return "x"
        if mode == CEILING_DIRECTION_Y:
            return "y"
        return "x" if span_x <= span_y else "y"

    def _member_point(self, ceiling_info, local_x, local_y, family_name, type_name):
        """Return the member centerline point for ceiling framing placement."""
        placement_mode = getattr(
            self.config,
            "ceiling_placement_mode",
            CEILING_PLACEMENT_ABOVE,
        )
        if placement_mode == CEILING_PLACEMENT_CENTER:
            return ceiling_info.point_at(local_x, local_y)

        member_depth = self._resolve_member_depth(family_name, type_name)
        depth_offset = -ceiling_info.target_layer_depth - (member_depth / 2.0)
        return ceiling_info.point_at(local_x, local_y, depth_offset)

    def _resolve_member_depth(self, family_name, type_name):
        """Resolve member depth from the family symbol or nominal lumber size."""
        depth = self.get_type_depth(family_name, type_name)
        if depth is not None and depth > 0.0:
            return depth

        text = "{0} {1}".format(family_name or "", type_name or "").lower()
        for nominal, dimensions in LUMBER_ACTUAL.items():
            if nominal.lower() in text:
                return inches_to_feet(dimensions[1])

        match = re.search(r"\b2x(2|3|4|6|8|10|12)\b", text)
        if match:
            nominal = "2x{0}".format(match.group(1))
            dims = LUMBER_ACTUAL.get(nominal)
            if dims is not None:
                return inches_to_feet(dims[1])

        return inches_to_feet(5.5)