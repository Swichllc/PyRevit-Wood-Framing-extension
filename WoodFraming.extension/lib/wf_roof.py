# -*- coding: utf-8 -*-
"""Roof framing engine -- stick-frame rafters and truss placement.

Supports gable, hip, shed/mono-slope, dutch-gable, and flat roofs.

Stick-frame construction sequence (real-world order):
  1. Ridge board(s) -- the beam at the peak where rafters meet
  2. Common rafters -- sloping members from ridge to wall plate at OC spacing
  3. Collar ties -- horizontal ties connecting opposing rafters (gable)
  4. Ceiling joists -- horizontal members eave-to-eave at plate elevation
  5. Kickers / outriggers -- diagonal braces from ceiling joist to rafter
  6. Sub-fascia -- board along the LOW eave edge (NOT rake edges)
  7. Ledger / header beam -- the high-side connection on shed roofs

Edge classification:
  - Ridge edges: shared boundary between two sloped faces (highest Z)
  - Eave edges: parallel to ridge at LOWER elevation -- gets fascia
  - Rake edges: perpendicular to ridge (gable-end slope edges)
  - Ledger edges: parallel to ridge at HIGHER elevation (shed roof high side)
"""

import math
import re

from wf_geometry import FramingMember, inches_to_feet
from wf_config import LUMBER_ACTUAL
from wf_host import (
    PlanarHostInfo,
    _extract_face_loops,
    _face_normal,
    _scanline_intervals,
    _to_local,
    analyze_roof_host,
)
from wf_placement import BaseFramingEngine
from wf_tracking import get_tracking_data


MIN_MEMBER_LENGTH = inches_to_feet(1.0)
RIDGE_TOL = inches_to_feet(3.0)
EDGE_TOL = inches_to_feet(1.0)
COLLAR_TIE_FRACTION = 1.0 / 3.0
MAX_COLLAR_TIE_SPACING = inches_to_feet(48.0)
PLATE_THICKNESS = inches_to_feet(1.5)
KICKER_FRACTION = 0.25
PROFILE_MATCH_TOL = inches_to_feet(0.125)

# Only faces within ~1 degree of horizontal are considered flat.
FLAT_THRESHOLD = 0.9998

# Edges whose direction dot product with the ridge exceeds this are
# considered "parallel" (eave/ledger).  Below this = rake.
PARALLEL_DOT_THRESHOLD = 0.5


# ======================================================================
#  Helpers
# ======================================================================

def _sloped_planes(planes):
    return [p for p in planes if p.normal.Z < FLAT_THRESHOLD]

def _classify_roof(planes):
    """Classify roof shape from its analyzed planes."""
    if not planes:
        return "flat"
    sloped = _sloped_planes(planes)
    flat_p = [p for p in planes if p.normal.Z >= FLAT_THRESHOLD]
    if not sloped:
        return "flat"
    if len(sloped) == 1:
        return "shed"
    if len(sloped) == 2 and not flat_p:
        return "gable"
    if len(sloped) == 4 and not flat_p:
        return "hip"
    if len(sloped) >= 2 and flat_p:
        return "dutch"
    return "complex"


def _single_slope_support_status(planes):
    """Return whether the current roof can use the stable single-slope path."""
    sloped = _sloped_planes(planes)
    roof_type = _classify_roof(planes)

    if len(sloped) == 1:
        return True, None, roof_type

    if not sloped:
        return (
            False,
            "Single-Slope Roof Framing requires exactly one sloped roof plane. "
            "This roof has no sloped planes.",
            roof_type,
        )

    return (
        False,
        "Single-Slope Roof Framing requires exactly one sloped roof plane. "
        "This roof has {0} sloped planes ({1}).".format(
            len(sloped),
            roof_type,
        ),
        roof_type,
    )


def _pt_key(pt, decimals=4):
    return (round(pt.X, decimals), round(pt.Y, decimals),
            round(pt.Z, decimals))


def _dist(a, b):
    return a.DistanceTo(b)


def _normalize(v):
    l = v.GetLength()
    if l < 1e-9:
        return None
    return v.Multiply(1.0 / l)


def _midpoint(a, b):
    from Autodesk.Revit.DB import XYZ
    return XYZ((a.X + b.X) / 2.0, (a.Y + b.Y) / 2.0, (a.Z + b.Z) / 2.0)


def _lerp(a, b, t):
    """Linear interpolation between two XYZ points."""
    from Autodesk.Revit.DB import XYZ
    return XYZ(
        a.X + t * (b.X - a.X),
        a.Y + t * (b.Y - a.Y),
        a.Z + t * (b.Z - a.Z),
    )


def _project_perpendicular(vector, axis):
    """Project a vector onto the plane perpendicular to the axis."""
    axis_unit = _normalize(axis)
    if axis_unit is None:
        return None
    return vector - axis_unit.Multiply(vector.DotProduct(axis_unit))


def _beam_reference_up(member_dir):
    """Return Revit's zero-rotation up vector for a beam-like member."""
    from Autodesk.Revit.DB import XYZ

    reference_up = _normalize(_project_perpendicular(XYZ.BasisZ, member_dir))
    if reference_up is not None:
        return reference_up
    return _normalize(_project_perpendicular(XYZ.BasisX, member_dir))


def _signed_angle_about(axis, start_vec, end_vec):
    """Return the signed angle from start_vec to end_vec about axis."""
    axis_unit = _normalize(axis)
    start_unit = _normalize(start_vec)
    end_unit = _normalize(end_vec)
    if axis_unit is None or start_unit is None or end_unit is None:
        return 0.0

    cross = start_unit.CrossProduct(end_unit)
    sin_value = axis_unit.DotProduct(cross)
    cos_value = max(-1.0, min(1.0, start_unit.DotProduct(end_unit)))
    return math.atan2(sin_value, cos_value)


def _rotation_from_up(member_dir, desired_up):
    """Convert a desired member up vector into Revit bend-direction rotation."""
    if desired_up is None:
        return 0.0

    reference_up = _beam_reference_up(member_dir)
    desired_up = _normalize(_project_perpendicular(desired_up, member_dir))
    if reference_up is None or desired_up is None:
        return 0.0
    return _signed_angle_about(member_dir, reference_up, desired_up)


# ======================================================================
#  Ridge / eave / rake detection
# ======================================================================

def _surface_point(plane, lx, ly):
    """World point on the face surface -- NO depth offset."""
    return (plane.origin
            + plane.x_axis.Multiply(lx)
            + plane.y_axis.Multiply(ly))


def _plane_point_at_depth(plane, lx, ly, depth_from_exterior):
    """World point on a plane parallel to the roof face at a given depth."""
    return (
        plane.origin
        + plane.x_axis.Multiply(lx)
        + plane.y_axis.Multiply(ly)
        + plane.normal.Multiply(-depth_from_exterior)
    )


def _points_near(first, second, tolerance=PROFILE_MATCH_TOL):
    """Check whether two XYZ points are within a small modeling tolerance."""
    try:
        return first.DistanceTo(second) <= tolerance
    except Exception:
        return False


def _same_segment(start_a, end_a, start_b, end_b, tolerance=PROFILE_MATCH_TOL):
    """Check whether two segments represent the same geometric edge."""
    return (
        (_points_near(start_a, start_b, tolerance)
         and _points_near(end_a, end_b, tolerance))
        or (_points_near(start_a, end_b, tolerance)
            and _points_near(end_a, start_b, tolerance))
    )


def _line_intersection_2d(line_a, line_b):
    """Return the intersection of two infinite 2D lines, or None."""
    (x1, y1), (x2, y2) = line_a
    (x3, y3), (x4, y4) = line_b

    denom = ((x1 - x2) * (y3 - y4)) - ((y1 - y2) * (x3 - x4))
    if abs(denom) < 1e-9:
        return None

    det_a = (x1 * y2) - (y1 * x2)
    det_b = (x3 * y4) - (y3 * x4)
    x = ((det_a * (x3 - x4)) - ((x1 - x2) * det_b)) / denom
    y = ((det_a * (y3 - y4)) - ((y1 - y2) * det_b)) / denom
    return (x, y)


def _segment_length(start_point, end_point):
    """Return segment length, or 0 on failure."""
    try:
        return (end_point - start_point).GetLength()
    except Exception:
        return 0.0


def _find_ridge_edges(planes):
    """Find ridge edges -- shared boundary between two sloped faces.

    Returns list of (start_xyz, end_xyz, plane_a, plane_b).
    """
    sloped = [p for p in planes if p.normal.Z < FLAT_THRESHOLD]
    if len(sloped) < 2:
        return []

    all_edges = []
    for plane in sloped:
        for loop in plane.boundary_loops_local:
            count = len(loop)
            for i in range(count):
                s_local = loop[i]
                e_local = loop[(i + 1) % count]
                s_world = _surface_point(plane, s_local[0], s_local[1])
                e_world = _surface_point(plane, e_local[0], e_local[1])
                all_edges.append((s_world, e_world, plane))

    ridges = []
    used = set()
    for i in range(len(all_edges)):
        if i in used:
            continue
        s1, e1, p1 = all_edges[i]
        for j in range(i + 1, len(all_edges)):
            if j in used:
                continue
            s2, e2, p2 = all_edges[j]
            if p1 is p2:
                continue
            match = False
            if _dist(s1, s2) < EDGE_TOL and _dist(e1, e2) < EDGE_TOL:
                match = True
            elif _dist(s1, e2) < EDGE_TOL and _dist(e1, s2) < EDGE_TOL:
                match = True
            if match:
                avg_z = (s1.Z + e1.Z) / 2.0
                ridges.append((s1, e1, p1, p2, avg_z))
                used.add(i)
                used.add(j)
                break

    if not ridges:
        return []

    # Keep all valid ridge segments (not only the top-most by Z).
    unique = []
    seen = set()
    for s, e, pa, pb, z in ridges:
        key = (_pt_key(s), _pt_key(e))
        rkey = (_pt_key(e), _pt_key(s))
        if key in seen or rkey in seen:
            continue
        seen.add(key)
        unique.append((s, e, pa, pb))

    unique.sort(
        key=lambda r: (((r[0].Z + r[1].Z) / 2.0), _dist(r[0], r[1])),
        reverse=True,
    )
    return unique


def _get_non_ridge_edges(plane, ridge_edges):
    """Return all boundary edges of a plane that are NOT ridge edges."""
    ridge_set = set()
    for rs, re, pa, pb in ridge_edges:
        ridge_set.add((_pt_key(rs), _pt_key(re)))
        ridge_set.add((_pt_key(re), _pt_key(rs)))

    edges = []
    for loop in plane.boundary_loops_local:
        count = len(loop)
        for i in range(count):
            s_local = loop[i]
            e_local = loop[(i + 1) % count]
            s_world = _surface_point(plane, s_local[0], s_local[1])
            e_world = _surface_point(plane, e_local[0], e_local[1])
            sk = _pt_key(s_world)
            ek = _pt_key(e_world)
            if (sk, ek) not in ridge_set:
                edges.append((s_world, e_world))
    return edges


def _classify_boundary_edges(plane, ridge_edges, ridge_segment=None):
    """Classify non-ridge boundary edges into eave, rake, and ledger.

    - ridge_dir: direction along the ridge (or plane.x_axis for sheds)
    - Eave: parallel to ridge AND at lower Z  --> fascia goes here
    - Ledger: parallel to ridge AND at higher Z  --> ledger beam (shed)
    - Rake: perpendicular to ridge  --> barge/rake board

    Returns dict with keys 'eave', 'rake', 'ledger', each a list of
    (start_xyz, end_xyz) tuples.
    """
    non_ridge = _get_non_ridge_edges(plane, ridge_edges)
    if not non_ridge:
        return {"eave": [], "rake": [], "ledger": []}

    # Determine ridge direction for this plane
    if ridge_segment is not None:
        rs, re = ridge_segment
        ridge_dir = _normalize(re - rs)
    else:
        plane_ridges = [(rs, re) for rs, re, pa, pb in ridge_edges
                        if pa is plane or pb is plane]
        if plane_ridges:
            plane_ridges.sort(
                key=lambda edge: _dist(edge[0], edge[1]),
                reverse=True,
            )
            rs, re = plane_ridges[0]
            ridge_dir = _normalize(re - rs)
        else:
            # Shed / no ridge: use plane x_axis (runs along the "ridge" direction)
            ridge_dir = _normalize(plane.x_axis)

    if ridge_dir is None:
        return {"eave": non_ridge, "rake": [], "ledger": []}

    # Separate parallel vs perpendicular edges
    parallel = []  # (start, end, avg_z)
    perpendicular = []
    for s, e in non_ridge:
        edge_dir = _normalize(e - s)
        if edge_dir is None:
            continue
        dot = abs(edge_dir.DotProduct(ridge_dir))
        avg_z = (s.Z + e.Z) / 2.0
        if dot >= PARALLEL_DOT_THRESHOLD:
            parallel.append((s, e, avg_z))
        else:
            perpendicular.append((s, e))

    # Split parallel edges into low (eave) and high (ledger)
    eave = []
    ledger = []
    if parallel:
        z_vals = [z for _, _, z in parallel]
        z_min = min(z_vals)
        z_max = max(z_vals)
        z_mid = (z_min + z_max) / 2.0

        if z_max - z_min < EDGE_TOL:
            # All at same elevation -- all are eaves (flat case)
            eave = [(s, e) for s, e, _ in parallel]
        else:
            for s, e, z in parallel:
                if z <= z_mid:
                    eave.append((s, e))
                else:
                    ledger.append((s, e))

    return {"eave": eave, "rake": perpendicular, "ledger": ledger}


def _classify_profile_boundary_edges(
    plane,
    loops_local,
    depth_from_exterior,
    ridge_edges=None,
    ridge_segment=None,
):
    """Classify boundary edges from an explicit roof profile loop set."""
    if not loops_local:
        return {"eave": [], "rake": [], "ledger": []}

    if ridge_segment is not None:
        rs, re = ridge_segment
        ridge_dir = _normalize(re - rs)
    else:
        ridge_dir = None
        if ridge_edges:
            plane_ridges = [(rs, re) for rs, re, pa, pb in ridge_edges
                            if pa is plane or pb is plane]
            if plane_ridges:
                plane_ridges.sort(
                    key=lambda edge: _dist(edge[0], edge[1]),
                    reverse=True,
                )
                rs, re = plane_ridges[0]
                ridge_dir = _normalize(re - rs)
        if ridge_dir is None:
            ridge_dir = _normalize(plane.x_axis)

    if ridge_dir is None:
        return {"eave": [], "rake": [], "ledger": []}

    parallel = []
    perpendicular = []
    for loop in loops_local:
        count = len(loop)
        for index in range(count):
            start_local = loop[index]
            end_local = loop[(index + 1) % count]
            start_world = _plane_point_at_depth(
                plane,
                start_local[0],
                start_local[1],
                depth_from_exterior,
            )
            end_world = _plane_point_at_depth(
                plane,
                end_local[0],
                end_local[1],
                depth_from_exterior,
            )
            edge_dir = _normalize(end_world - start_world)
            if edge_dir is None:
                continue
            dot = abs(edge_dir.DotProduct(ridge_dir))
            avg_z = (start_world.Z + end_world.Z) / 2.0
            if dot >= PARALLEL_DOT_THRESHOLD:
                parallel.append((start_world, end_world, avg_z))
            else:
                perpendicular.append((start_world, end_world))

    eave = []
    ledger = []
    if parallel:
        z_vals = [z for _, _, z in parallel]
        z_min = min(z_vals)
        z_max = max(z_vals)
        z_mid = (z_min + z_max) / 2.0

        if z_max - z_min < EDGE_TOL:
            eave = [(s, e) for s, e, _ in parallel]
        else:
            for s, e, z in parallel:
                if z <= z_mid:
                    eave.append((s, e))
                else:
                    ledger.append((s, e))

    return {"eave": eave, "rake": perpendicular, "ledger": ledger}


def _lowest_eave_z(eave_edges):
    """Return the lowest Z among all eave edge endpoints."""
    if not eave_edges:
        return None
    z_vals = []
    for s, e in eave_edges:
        z_vals.append(s.Z)
        z_vals.append(e.Z)
    return min(z_vals) if z_vals else None


def _rafter_positions(eave_start, eave_end, spacing):
    """Generate OC positions along an eave edge."""
    edge_len = _dist(eave_start, eave_end)
    if edge_len < MIN_MEMBER_LENGTH or spacing <= 0:
        return []
    positions = []
    d = 0.0
    while d <= edge_len + 1e-9:
        t = min(d / edge_len, 1.0)
        positions.append(t)
        d += spacing
    if abs(positions[-1] - 1.0) > 1e-6:
        positions.append(1.0)
    return positions


def _project_to_ridge(pt, ridge_start, ridge_end):
    """Project a point onto a ridge line."""
    ridge_dir = ridge_end - ridge_start
    ridge_len = ridge_dir.GetLength()
    if ridge_len < 1e-9:
        return ridge_start
    ridge_dir = ridge_dir.Multiply(1.0 / ridge_len)
    t = (pt - ridge_start).DotProduct(ridge_dir)
    t = max(0.0, min(ridge_len, t))
    return ridge_start + ridge_dir.Multiply(t)


def _project_to_edge(pt, edge_start, edge_end):
    """Project a point onto an edge segment and report whether clamping was needed."""
    edge_dir = edge_end - edge_start
    edge_len = edge_dir.GetLength()
    if edge_len < 1e-9:
        return edge_start, False, 0.0

    edge_unit = edge_dir.Multiply(1.0 / edge_len)
    raw_t = (pt - edge_start).DotProduct(edge_unit)
    clamped_t = max(0.0, min(edge_len, raw_t))
    was_clamped = abs(clamped_t - raw_t) > 1e-9
    return edge_start + edge_unit.Multiply(clamped_t), (not was_clamped), raw_t


def _ridge_station_on_segment(point, ridge_start, ridge_end):
    """Return the station of a point along a ridge segment."""
    ridge_dir = ridge_end - ridge_start
    ridge_len = ridge_dir.GetLength()
    if ridge_len < 1e-9:
        return 0.0
    ridge_unit = ridge_dir.Multiply(1.0 / ridge_len)
    return (point - ridge_start).DotProduct(ridge_unit)


def _segment_covers_ridge_station(edge_start, edge_end, ridge_start, ridge_end, ridge_station):
    """Return True when an edge spans the current ridge station."""
    edge_station_start = _ridge_station_on_segment(edge_start, ridge_start, ridge_end)
    edge_station_end = _ridge_station_on_segment(edge_end, ridge_start, ridge_end)
    station_min = min(edge_station_start, edge_station_end)
    station_max = max(edge_station_start, edge_station_end)
    return (station_min - EDGE_TOL) <= ridge_station <= (station_max + EDGE_TOL)


def _project_to_best_eave(ridge_pt, eave_edges, ridge_start=None, ridge_end=None):
    """Project a ridge point to the nearest valid eave edge."""
    best_pt = None
    best_covers_station = False
    best_interior = False
    best_dist = None
    ridge_station = None
    if ridge_start is not None and ridge_end is not None:
        ridge_station = _ridge_station_on_segment(ridge_pt, ridge_start, ridge_end)
    for eave_s, eave_e in eave_edges:
        cand, is_interior, _ = _project_to_edge(ridge_pt, eave_s, eave_e)
        covers_station = False
        if ridge_station is not None:
            covers_station = _segment_covers_ridge_station(
                eave_s,
                eave_e,
                ridge_start,
                ridge_end,
                ridge_station,
            )
        d = _dist(ridge_pt, cand)
        if best_pt is None:
            best_pt = cand
            best_covers_station = covers_station
            best_interior = is_interior
            best_dist = d
            continue

        if covers_station and not best_covers_station:
            best_pt = cand
            best_covers_station = True
            best_interior = is_interior
            best_dist = d
            continue

        if is_interior and not best_interior:
            best_pt = cand
            best_covers_station = covers_station
            best_interior = True
            best_dist = d
            continue

        if (covers_station == best_covers_station
                and is_interior == best_interior
                and d < best_dist):
            best_pt = cand
            best_covers_station = covers_station
            best_interior = is_interior
            best_dist = d
    return best_pt


# ======================================================================
#  Engine
# ======================================================================

class RoofFramingEngine(BaseFramingEngine):
    """Calculates and places roof framing -- stick-frame or truss."""

    def place_members(self, members, host_info):
        """Place roof members and apply roof-specific post-processing."""
        placed = BaseFramingEngine.place_members(self, members, host_info)

        try:
            self._set_coping_distance_zero(placed)
        except Exception:
            pass

        if placed:
            try:
                self.doc.Regenerate()
            except Exception:
                pass

        try:
            self._apply_automatic_coping(placed)
        except Exception:
            pass

        return placed

    @staticmethod
    def _set_coping_distance_zero(instances):
        """Set coping distance to zero on roof framing when the parameter exists."""
        try:
            from Autodesk.Revit.DB import BuiltInParameter
        except Exception:
            BuiltInParameter = None

        if BuiltInParameter is None:
            return

        for instance in instances or []:
            try:
                parameter = instance.get_Parameter(
                    BuiltInParameter.STRUCTURAL_COPING_DISTANCE,
                )
            except Exception:
                parameter = None
            if parameter is None:
                continue
            try:
                if not parameter.IsReadOnly:
                    parameter.Set(0.0)
            except Exception:
                pass

    def calculate_members(self, roof, mode="stick"):
        """Calculate framing members for a roof.

        Args:
            roof: Revit RoofBase element.
            mode: "stick" for rafter framing, "truss" for truss placement.

        Returns:
            (members_list, roof_info) or ([], None) on failure.
        """
        roof_info = analyze_roof_host(self.doc, roof, self.config)
        if roof_info is None:
            return [], None
        is_supported, support_reason, roof_type = _single_slope_support_status(
            getattr(roof_info, "planes", []) or []
        )
        try:
            roof_info.roof_type = roof_type
            roof_info.single_slope_supported = is_supported
            roof_info.single_slope_support_reason = support_reason
        except Exception:
            pass
        if not is_supported:
            return [], roof_info
        if mode == "truss":
            members = self._calc_truss_positions(roof_info)
        else:
            members = self._calc_stick_frame(roof_info)
        return members, roof_info

    def _apply_automatic_coping(self, placed_instances):
        """Best-effort coping between newly placed rafters and perimeter boards."""
        if not placed_instances:
            return

        rafters = []
        boards = []
        member_pairs = getattr(self, "_last_placed_pairs", None) or []

        if member_pairs:
            for member, instance in member_pairs:
                member_type = getattr(member, "member_type", None)
                if member_type == "RAFTER":
                    rafters.append(instance)
                elif member_type in ("FASCIA", "LEDGER"):
                    boards.append(instance)
        else:
            for instance in placed_instances:
                tracking = get_tracking_data(instance)
                if tracking is None:
                    continue
                member_type = tracking.get("member")
                if member_type == "RAFTER":
                    rafters.append(instance)
                elif member_type in ("FASCIA", "LEDGER"):
                    boards.append(instance)

        if not rafters or not boards:
            return

        for rafter in rafters:
            add_coping = getattr(rafter, "AddCoping", None)
            if add_coping is None:
                continue
            for board in boards:
                if not self._elements_are_near(rafter, board):
                    continue
                try:
                    add_coping(board)
                except Exception:
                    pass

    @staticmethod
    def _elements_are_near(first, second, tolerance=0.25):
        """Return True when two elements' bounding boxes overlap or nearly touch."""
        first_box = first.get_BoundingBox(None)
        second_box = second.get_BoundingBox(None)
        if first_box is None or second_box is None:
            return False

        return not (
            (first_box.Max.X + tolerance) < second_box.Min.X
            or (second_box.Max.X + tolerance) < first_box.Min.X
            or (first_box.Max.Y + tolerance) < second_box.Min.Y
            or (second_box.Max.Y + tolerance) < first_box.Min.Y
            or (first_box.Max.Z + tolerance) < second_box.Min.Z
            or (second_box.Max.Z + tolerance) < first_box.Min.Z
        )

    # ------------------------------------------------------------------
    #  Stick framing
    # ------------------------------------------------------------------

    def _calc_stick_frame(self, roof_info):
        members = []
        planes = roof_info.planes
        roof_type = _classify_roof(planes)
        spacing = self.config.stud_spacing_ft
        if spacing <= 0:
            return members

        # 1. Ridge detection
        try:
            ridge_edges = _find_ridge_edges(planes)
        except Exception:
            ridge_edges = []

        members.extend(self._make_ridge_boards(ridge_edges, roof_info))

        # 2. Rafters per slope + classify edges
        all_eave_edges = []
        all_rake_edges = []
        all_ledger_edges = []
        for plane in planes:
            if plane.normal.Z >= FLAT_THRESHOLD:
                continue

            # Classify this plane's boundary edges
            try:
                edge_depth = 0.0
                classified = _classify_boundary_edges(plane, ridge_edges)
                if roof_type == "shed":
                    edge_depth = self._resolve_roof_layer_top_depth(plane)
                    profile_loops = self._roof_profile_loops_local(
                        plane,
                        edge_depth,
                    )
                    if profile_loops:
                        classified = _classify_profile_boundary_edges(
                            plane,
                            profile_loops,
                            edge_depth,
                            ridge_edges,
                        )
                all_eave_edges.extend(
                    (edge_start, edge_end, plane, "eave", edge_depth)
                    for edge_start, edge_end in classified["eave"]
                )
                all_rake_edges.extend(
                    (edge_start, edge_end, plane, "rake", edge_depth)
                    for edge_start, edge_end in classified["rake"]
                )
                all_ledger_edges.extend(
                    (edge_start, edge_end, plane, "ledger", edge_depth)
                    for edge_start, edge_end in classified["ledger"]
                )
            except Exception:
                pass

            # Place rafters: ridged roofs prefer ridge/eave-controlled axes.
            rafters = []
            if ridge_edges and roof_type != "shed":
                try:
                    rafters = self._make_rafters_for_plane(
                        plane, ridge_edges, roof_info)
                except Exception:
                    rafters = []
                if not rafters:
                    try:
                        rafters = self._make_rafters_scanline(
                            plane, spacing, roof_info)
                    except Exception:
                        rafters = []
            else:
                try:
                    rafters = self._make_rafters_scanline(
                        plane, spacing, roof_info)
                except Exception:
                    rafters = []
                if not rafters and ridge_edges:
                    try:
                        rafters = self._make_rafters_for_plane(
                            plane, ridge_edges, roof_info)
                    except Exception:
                        rafters = []
            members.extend(rafters)

        # 3. Collar ties per ridge segment.
        if (
            ridge_edges
            and bool(getattr(self.config, "include_collar_ties", True))
        ):
            try:
                members.extend(
                    self._make_collar_ties(planes, ridge_edges, roof_info))
            except Exception:
                pass

        # 4. Ceiling joists + kickers per ridge segment.
        if (
            ridge_edges
            and bool(getattr(self.config, "include_ceiling_joists", True))
        ):
            try:
                members.extend(
                    self._make_ceiling_joists(
                        planes,
                        ridge_edges,
                        roof_info,
                        spacing,
                        bool(getattr(self.config, "include_roof_kickers", True)),
                    )
                )
            except Exception:
                pass

        # 5. Shed roofs need border members along the low eave and rakes.
        fascia_edges = list(all_eave_edges)
        if roof_type == "shed":
            fascia_edges.extend(all_rake_edges)
        try:
            members.extend(
                self._make_fascia(fascia_edges, roof_info))
        except Exception:
            pass

        # 6. Ledger beam at high side (shed roofs)
        if roof_type == "shed" and all_ledger_edges:
            try:
                members.extend(
                    self._make_ledger(all_ledger_edges, roof_info))
            except Exception:
                pass

        return members

    # ------------------------------------------------------------------
    #  Ridge boards
    # ------------------------------------------------------------------

    def _make_ridge_boards(self, ridge_edges, roof_info):
        members = []
        seen = set()
        for rs, re, pa, pb in ridge_edges:
            if _dist(rs, re) < MIN_MEMBER_LENGTH:
                continue
            key = (_pt_key(rs), _pt_key(re))
            rkey = (_pt_key(re), _pt_key(rs))
            if key in seen or rkey in seen:
                continue
            seen.add(key)
            m = FramingMember(FramingMember.HEADER, rs, re)
            m.member_type = "RIDGE_BOARD"
            m.family_name = (
                self.config.header_family_name or self.config.stud_family_name)
            m.type_name = (
                self.config.header_type_name or self.config.stud_type_name)
            m.rotation = 0.0
            m.host_kind = roof_info.kind
            m.host_id = roof_info.element_id
            members.append(m)
        return members

    # ------------------------------------------------------------------
    #  Rafters
    # ------------------------------------------------------------------

    def _make_rafters_for_plane(self, plane, ridge_edges, roof_info):
        members = []
        spacing = self.config.stud_spacing_ft
        if spacing <= 0:
            return members

        classified = _classify_boundary_edges(plane, ridge_edges)
        eave_edges = classified["eave"]
        if not eave_edges:
            return self._make_rafters_scanline(plane, spacing, roof_info)

        plane_ridges = [(rs, re) for rs, re, pa, pb in ridge_edges
                        if pa is plane or pb is plane]
        if not plane_ridges:
            return self._make_rafters_scanline(plane, spacing, roof_info)

        plane_ridges.sort(key=lambda edge: _dist(edge[0], edge[1]), reverse=True)
        ridge_s, ridge_e = plane_ridges[0]
        eave_edges.sort(key=lambda e: _dist(e[0], e[1]), reverse=True)

        seen = set()
        for eave_s, eave_e in eave_edges:
            eave_dir = eave_e - eave_s
            eave_len = eave_dir.GetLength()
            if eave_len < MIN_MEMBER_LENGTH:
                continue
            eave_dir = eave_dir.Multiply(1.0 / eave_len)

            positions = _rafter_positions(eave_s, eave_e, spacing)
            for t in positions:
                eave_pt = eave_s + eave_dir.Multiply(t * eave_len)
                ridge_pt = _project_to_ridge(eave_pt, ridge_s, ridge_e)
                if _dist(eave_pt, ridge_pt) < MIN_MEMBER_LENGTH:
                    continue
                key = (_pt_key(eave_pt), _pt_key(ridge_pt))
                rkey = (_pt_key(ridge_pt), _pt_key(eave_pt))
                if key in seen or rkey in seen:
                    continue
                seen.add(key)

                m = FramingMember(FramingMember.STUD, eave_pt, ridge_pt)
                m.member_type = "RAFTER"
                m.family_name = self.config.stud_family_name
                m.type_name = self.config.stud_type_name
                m.rotation = _rotation_from_up(ridge_pt - eave_pt, plane.normal)
                m.disallow_end_joins = True
                m.host_kind = plane.kind
                m.host_id = plane.element_id
                members.append(m)

        if not members:
            return self._make_rafters_scanline(plane, spacing, roof_info)
        return members

    def _make_rafters_scanline(self, plane, spacing, roof_info):
        """Fallback for shed / flat roofs with no ridge."""
        members = []
        family_name = self.config.stud_family_name
        type_name = self.config.stud_type_name
        control_depth = self._resolve_roof_member_center_depth(
            plane,
            family_name,
            type_name,
        )
        profile_loops = self._roof_profile_loops_local(plane, control_depth)
        if not profile_loops:
            return members

        points = [point for loop in profile_loops for point in loop]
        if not points:
            return members

        min_x = min(point[0] for point in points)
        max_x = max(point[0] for point in points)
        min_y = min(point[1] for point in points)
        max_y = max(point[1] for point in points)
        if max_y - min_y < MIN_MEMBER_LENGTH:
            return members

        x = min_x
        while x <= max_x + 1e-9:
            intervals = _scanline_intervals(profile_loops, "x", x)
            for start_y, end_y in intervals:
                if end_y - start_y < MIN_MEMBER_LENGTH:
                    continue
                s_pt = _plane_point_at_depth(plane, x, start_y, control_depth)
                e_pt = _plane_point_at_depth(plane, x, end_y, control_depth)
                original_length = _dist(s_pt, e_pt)
                clipped = self._clip_member_axis_to_roof(plane.element, s_pt, e_pt)
                if clipped is not None:
                    clipped_length = _dist(clipped[0], clipped[1])
                    if clipped_length <= original_length + 1e-6:
                        s_pt, e_pt = clipped
                if _dist(s_pt, e_pt) < MIN_MEMBER_LENGTH:
                    continue
                m = FramingMember(FramingMember.STUD, s_pt, e_pt)
                m.member_type = "RAFTER"
                m.family_name = self.config.stud_family_name
                m.type_name = self.config.stud_type_name
                m.rotation = _rotation_from_up(e_pt - s_pt, plane.normal)
                m.disallow_end_joins = True
                m.host_kind = plane.kind
                m.host_id = plane.element_id
                members.append(m)
            x += spacing
        return members

    def _resolve_roof_member_center_depth(self, plane, family_name, type_name):
        """Return the control-plane depth for a roof framing member centerline."""
        layer_top_depth = self._resolve_roof_layer_top_depth(plane)
        _, member_depth = self._resolve_roof_member_size(family_name, type_name)
        if member_depth > 0.0:
            return layer_top_depth + (member_depth / 2.0)

        target_layer_depth = getattr(plane, "target_layer_depth", 0.0)
        if target_layer_depth > 0.0:
            return target_layer_depth
        return layer_top_depth

    def _roof_profile_loops_local(self, plane, depth_from_exterior):
        """Derive a roof member control profile from adjacent roof faces."""
        if depth_from_exterior <= 1e-9:
            return plane.boundary_loops_local

        adjacent_faces = self._collect_adjacent_roof_faces(plane.element, plane.normal)
        if not adjacent_faces:
            return plane.boundary_loops_local

        shifted_loops = []
        for loop in plane.boundary_loops_local:
            world_loop = [
                _surface_point(plane, local_x, local_y)
                for local_x, local_y in loop
            ]
            shifted_loop = self._shift_roof_loop_local(
                plane,
                world_loop,
                depth_from_exterior,
                adjacent_faces,
            )
            if shifted_loop is None:
                return plane.boundary_loops_local
            shifted_loops.append(shifted_loop)

        return shifted_loops

    def _collect_adjacent_roof_faces(self, roof, top_normal):
        """Collect non-coplanar roof faces that can bound the framing profile."""
        from Autodesk.Revit.DB import GeometryInstance, Options, Solid, ViewDetailLevel

        try:
            options = Options()
            options.ComputeReferences = False
            options.DetailLevel = ViewDetailLevel.Fine
            geometry = roof.get_Geometry(options)
        except Exception:
            geometry = None
        if geometry is None:
            return []

        solids = []
        for geom_obj in geometry:
            if isinstance(geom_obj, Solid) and geom_obj.Volume > 0:
                solids.append(geom_obj)
                continue
            if isinstance(geom_obj, GeometryInstance):
                try:
                    instance_geometry = geom_obj.GetInstanceGeometry()
                except Exception:
                    instance_geometry = None
                if instance_geometry is None:
                    continue
                for sub_obj in instance_geometry:
                    if isinstance(sub_obj, Solid) and sub_obj.Volume > 0:
                        solids.append(sub_obj)

        adjacent_faces = []
        for solid in solids:
            for face in solid.Faces:
                face_normal = _face_normal(face)
                if face_normal is None:
                    continue
                if abs(face_normal.DotProduct(top_normal)) > 0.9999:
                    continue
                face_loops = _extract_face_loops(face)
                if face_loops:
                    adjacent_faces.append((face_normal, face_loops))

        return adjacent_faces

    def _shift_roof_loop_local(self, plane, loop_points, depth_from_exterior, adjacent_faces):
        """Project a top-face loop down to a parallel control plane via adjacent faces."""
        if len(loop_points) < 3:
            return None

        shifted_lines = []
        for index in range(len(loop_points)):
            start_point = loop_points[index]
            end_point = loop_points[(index + 1) % len(loop_points)]
            edge_dir = _normalize(end_point - start_point)
            if edge_dir is None:
                return None

            face_normal = self._find_adjacent_face_normal(
                adjacent_faces,
                start_point,
                end_point,
            )
            if face_normal is None:
                return None

            move_axis = _normalize(edge_dir.CrossProduct(face_normal))
            denominator = plane.normal.DotProduct(move_axis) if move_axis is not None else 0.0
            if move_axis is None or abs(denominator) < 1e-9:
                move_axis = _normalize(face_normal.CrossProduct(edge_dir))
                denominator = plane.normal.DotProduct(move_axis) if move_axis is not None else 0.0
            if move_axis is None or abs(denominator) < 1e-9:
                return None

            distance = -depth_from_exterior / denominator
            shift_vec = move_axis.Multiply(distance)
            shift_x = shift_vec.DotProduct(plane.x_axis)
            shift_y = shift_vec.DotProduct(plane.y_axis)

            start_local = _to_local(start_point, plane.origin, plane.x_axis, plane.y_axis)
            end_local = _to_local(end_point, plane.origin, plane.x_axis, plane.y_axis)
            shifted_lines.append(
                (
                    (start_local[0] + shift_x, start_local[1] + shift_y),
                    (end_local[0] + shift_x, end_local[1] + shift_y),
                )
            )

        shifted_loop = []
        for index in range(len(shifted_lines)):
            prev_line = shifted_lines[index - 1]
            curr_line = shifted_lines[index]
            point = _line_intersection_2d(prev_line, curr_line)
            if point is None:
                point = curr_line[0]
            shifted_loop.append(point)

        return shifted_loop

    def _find_adjacent_face_normal(self, adjacent_faces, start_point, end_point):
        """Find the non-coplanar roof face that shares a boundary segment."""
        for face_normal, loops in adjacent_faces:
            for loop in loops:
                count = len(loop)
                for index in range(count):
                    face_start = loop[index]
                    face_end = loop[(index + 1) % count]
                    if _same_segment(start_point, end_point, face_start, face_end):
                        return face_normal
        return None

    def _get_roof_solids(self, roof):
        """Return cached positive-volume solids for a roof element."""
        from Autodesk.Revit.DB import GeometryInstance, Options, Solid, ViewDetailLevel

        cache = getattr(self, "_roof_solids_cache", None)
        if cache is None:
            cache = {}
            self._roof_solids_cache = cache

        key = getattr(getattr(roof, "Id", None), "IntegerValue", None)
        if key in cache:
            return cache[key]

        solids = []
        try:
            options = Options()
            options.ComputeReferences = False
            options.DetailLevel = ViewDetailLevel.Fine
            geometry = roof.get_Geometry(options)
        except Exception:
            geometry = None

        if geometry is not None:
            for geom_obj in geometry:
                if isinstance(geom_obj, Solid) and geom_obj.Volume > 0:
                    solids.append(geom_obj)
                    continue
                if isinstance(geom_obj, GeometryInstance):
                    try:
                        instance_geometry = geom_obj.GetInstanceGeometry()
                    except Exception:
                        instance_geometry = None
                    if instance_geometry is None:
                        continue
                    for sub_obj in instance_geometry:
                        if isinstance(sub_obj, Solid) and sub_obj.Volume > 0:
                            solids.append(sub_obj)

        cache[key] = solids
        return solids

    def _clip_member_axis_to_roof(self, roof, start_point, end_point):
        """Clip a member axis to the actual roof solid along its line of action."""
        from Autodesk.Revit.DB import Line, SolidCurveIntersectionOptions

        axis_dir = _normalize(end_point - start_point)
        axis_len = _segment_length(start_point, end_point)
        if axis_dir is None or axis_len < MIN_MEMBER_LENGTH:
            return None

        solids = self._get_roof_solids(roof)
        if not solids:
            return (start_point, end_point)

        try:
            bbox = roof.get_BoundingBox(None)
        except Exception:
            bbox = None

        probe_half = axis_len + 10.0
        if bbox is not None:
            try:
                dx = bbox.Max.X - bbox.Min.X
                dy = bbox.Max.Y - bbox.Min.Y
                dz = bbox.Max.Z - bbox.Min.Z
                probe_half = max(probe_half, math.sqrt(dx * dx + dy * dy + dz * dz) + 10.0)
            except Exception:
                pass

        mid_point = _midpoint(start_point, end_point)
        probe_start = mid_point - axis_dir.Multiply(probe_half)
        probe_end = mid_point + axis_dir.Multiply(probe_half)

        try:
            probe_line = Line.CreateBound(probe_start, probe_end)
        except Exception:
            return (start_point, end_point)

        target_t = (mid_point - probe_start).DotProduct(axis_dir)
        best_segment = None
        best_contains_target = False
        best_distance = None
        best_length = 0.0

        for solid in solids:
            try:
                result = solid.IntersectWithCurve(
                    probe_line,
                    SolidCurveIntersectionOptions(),
                )
                seg_count = result.SegmentCount
            except Exception:
                continue

            for index in range(seg_count):
                try:
                    segment = result.GetCurveSegment(index)
                    seg_start = segment.GetEndPoint(0)
                    seg_end = segment.GetEndPoint(1)
                except Exception:
                    continue

                start_t = (seg_start - probe_start).DotProduct(axis_dir)
                end_t = (seg_end - probe_start).DotProduct(axis_dir)
                seg_min = min(start_t, end_t)
                seg_max = max(start_t, end_t)
                contains_target = (seg_min - 1e-6) <= target_t <= (seg_max + 1e-6)
                distance = 0.0 if contains_target else min(abs(target_t - seg_min), abs(target_t - seg_max))
                seg_length = _segment_length(seg_start, seg_end)

                choose = False
                if best_segment is None:
                    choose = True
                elif contains_target and not best_contains_target:
                    choose = True
                elif contains_target == best_contains_target:
                    if distance < (best_distance if best_distance is not None else float("inf")) - 1e-6:
                        choose = True
                    elif abs(distance - (best_distance if best_distance is not None else distance)) <= 1e-6 and seg_length > best_length:
                        choose = True

                if choose:
                    if start_t <= end_t:
                        best_segment = (seg_start, seg_end)
                    else:
                        best_segment = (seg_end, seg_start)
                    best_contains_target = contains_target
                    best_distance = distance
                    best_length = seg_length

        return best_segment

    def _choose_best_side_shift(self, roof, start_point, end_point, side_axis, distance):
        """Pick the side shift whose clipped axis stays farther inside the roof."""
        if side_axis is None or distance <= 1e-9:
            return None

        best_segment = None
        best_length = -1.0
        for sign in (-1.0, 1.0):
            shift = side_axis.Multiply(sign * distance)
            cand_start = start_point + shift
            cand_end = end_point + shift
            clipped = self._clip_member_axis_to_roof(roof, cand_start, cand_end)
            if clipped is None:
                continue
            seg_length = _dist(clipped[0], clipped[1])
            if seg_length > best_length + 1e-6:
                best_segment = clipped
                best_length = seg_length

        return best_segment

    # ------------------------------------------------------------------
    #  Collar ties
    # ------------------------------------------------------------------

    def _make_collar_ties(self, planes, ridge_edges, roof_info):
        """Collar ties at 1/3 rafter length from ridge, every other rafter."""
        from Autodesk.Revit.DB import XYZ

        members = []
        seen = set()
        if not ridge_edges:
            return members

        spacing = self.config.stud_spacing_ft
        if spacing <= 0:
            return members

        tie_spacing = spacing * 2.0
        if tie_spacing > MAX_COLLAR_TIE_SPACING:
            tie_spacing = spacing
        if tie_spacing > MAX_COLLAR_TIE_SPACING:
            tie_spacing = MAX_COLLAR_TIE_SPACING

        for rs, re, pa, pb in ridge_edges:
            ridge_dir = re - rs
            ridge_len = ridge_dir.GetLength()
            if ridge_len < MIN_MEMBER_LENGTH:
                continue
            ridge_unit = ridge_dir.Multiply(1.0 / ridge_len)

            class_a = _classify_boundary_edges(pa, ridge_edges, (rs, re))
            class_b = _classify_boundary_edges(pb, ridge_edges, (rs, re))
            eave_a = class_a["eave"]
            eave_b = class_b["eave"]
            if not eave_a or not eave_b:
                continue

            d = tie_spacing / 2.0
            while d < ridge_len:
                ridge_pt = rs + ridge_unit.Multiply(d)
                foot_a = _project_to_best_eave(ridge_pt, eave_a, rs, re)
                foot_b = _project_to_best_eave(ridge_pt, eave_b, rs, re)
                if foot_a is None or foot_b is None:
                    d += tie_spacing
                    continue

                tie_a = _lerp(ridge_pt, foot_a, COLLAR_TIE_FRACTION)
                tie_b = _lerp(ridge_pt, foot_b, COLLAR_TIE_FRACTION)
                tie_z = (tie_a.Z + tie_b.Z) / 2.0
                tie_a = XYZ(tie_a.X, tie_a.Y, tie_z)
                tie_b = XYZ(tie_b.X, tie_b.Y, tie_z)

                if _dist(tie_a, tie_b) >= MIN_MEMBER_LENGTH:
                    key = (_pt_key(tie_a), _pt_key(tie_b))
                    rkey = (_pt_key(tie_b), _pt_key(tie_a))
                    if key not in seen and rkey not in seen:
                        seen.add(key)
                        m = FramingMember(FramingMember.HEADER, tie_a, tie_b)
                        m.member_type = "COLLAR_TIE"
                        m.family_name = self.config.stud_family_name
                        m.type_name = self.config.stud_type_name
                        m.rotation = -math.pi / 2.0  # flat
                        m.host_kind = roof_info.kind
                        m.host_id = roof_info.element_id
                        members.append(m)
                d += tie_spacing

        return members

    # ------------------------------------------------------------------
    #  Ceiling joists + kickers
    # ------------------------------------------------------------------

    def _make_ceiling_joists(self, planes, ridge_edges, roof_info, spacing, include_kickers=True):
        """Ceiling joists spanning eave-to-eave at plate line elevation.

        Also generates kicker/outrigger braces from each joist up to
        the rafter at KICKER_FRACTION of rafter length from eave.
        """
        from Autodesk.Revit.DB import XYZ

        members = []
        joist_seen = set()
        kicker_seen = set()
        if not ridge_edges:
            return members
        if spacing <= 0:
            return members

        for rs, re, pa, pb in ridge_edges:
            ridge_dir = re - rs
            ridge_len = ridge_dir.GetLength()
            if ridge_len < MIN_MEMBER_LENGTH:
                continue
            ridge_unit = ridge_dir.Multiply(1.0 / ridge_len)

            class_a = _classify_boundary_edges(pa, ridge_edges, (rs, re))
            class_b = _classify_boundary_edges(pb, ridge_edges, (rs, re))
            eave_a = class_a["eave"]
            eave_b = class_b["eave"]
            if not eave_a or not eave_b:
                continue

            d = 0.0
            while d <= ridge_len + 1e-9:
                ridge_pt = rs + ridge_unit.Multiply(min(d, ridge_len))
                foot_a = _project_to_best_eave(ridge_pt, eave_a, rs, re)
                foot_b = _project_to_best_eave(ridge_pt, eave_b, rs, re)
                if foot_a is None or foot_b is None:
                    d += spacing
                    continue

                # Use lower support point to avoid floating joists on uneven eaves.
                joist_z = min(foot_a.Z, foot_b.Z)
                joist_a = XYZ(foot_a.X, foot_a.Y, joist_z)
                joist_b = XYZ(foot_b.X, foot_b.Y, joist_z)

                if _dist(joist_a, joist_b) >= MIN_MEMBER_LENGTH:
                    jkey = (_pt_key(joist_a), _pt_key(joist_b))
                    jrkey = (_pt_key(joist_b), _pt_key(joist_a))
                    if jkey not in joist_seen and jrkey not in joist_seen:
                        joist_seen.add(jkey)
                        m = FramingMember(FramingMember.HEADER, joist_a, joist_b)
                        m.member_type = "CEILING_JOIST"
                        m.family_name = self.config.stud_family_name
                        m.type_name = self.config.stud_type_name
                        m.rotation = -math.pi / 2.0  # flat
                        m.host_kind = roof_info.kind
                        m.host_id = roof_info.element_id
                        members.append(m)

                    if include_kickers:
                        # Kicker side A: diagonal from joist to rafter.
                        rafter_pt_a = _lerp(foot_a, ridge_pt, KICKER_FRACTION)
                        kick_base_a = _lerp(joist_a, joist_b, KICKER_FRACTION)
                        if _dist(kick_base_a, rafter_pt_a) >= MIN_MEMBER_LENGTH:
                            kkey = (_pt_key(kick_base_a), _pt_key(rafter_pt_a))
                            krkey = (_pt_key(rafter_pt_a), _pt_key(kick_base_a))
                            if kkey not in kicker_seen and krkey not in kicker_seen:
                                kicker_seen.add(kkey)
                                km = FramingMember(
                                    FramingMember.STUD,
                                    kick_base_a,
                                    rafter_pt_a,
                                )
                                km.member_type = "KICKER"
                                km.family_name = self.config.stud_family_name
                                km.type_name = self.config.stud_type_name
                                km.rotation = _rotation_from_up(
                                    rafter_pt_a - kick_base_a,
                                    getattr(pa, "normal", None),
                                )
                                km.host_kind = roof_info.kind
                                km.host_id = roof_info.element_id
                                members.append(km)

                        # Kicker side B.
                        rafter_pt_b = _lerp(foot_b, ridge_pt, KICKER_FRACTION)
                        kick_base_b = _lerp(joist_b, joist_a, KICKER_FRACTION)
                        if _dist(kick_base_b, rafter_pt_b) >= MIN_MEMBER_LENGTH:
                            kkey = (_pt_key(kick_base_b), _pt_key(rafter_pt_b))
                            krkey = (_pt_key(rafter_pt_b), _pt_key(kick_base_b))
                            if kkey not in kicker_seen and krkey not in kicker_seen:
                                kicker_seen.add(kkey)
                                km = FramingMember(
                                    FramingMember.STUD,
                                    kick_base_b,
                                    rafter_pt_b,
                                )
                                km.member_type = "KICKER"
                                km.family_name = self.config.stud_family_name
                                km.type_name = self.config.stud_type_name
                                km.rotation = _rotation_from_up(
                                    rafter_pt_b - kick_base_b,
                                    getattr(pb, "normal", None),
                                )
                                km.host_kind = roof_info.kind
                                km.host_id = roof_info.element_id
                                members.append(km)

                d += spacing

        return members

    # ------------------------------------------------------------------
    #  Fascia / border trim
    # ------------------------------------------------------------------

    def _make_fascia(self, eave_edges, roof_info):
        """Create border trim members along the supplied roof boundary edges."""
        members = []
        seen = set()
        for es, ee, plane, edge_role, edge_depth in eave_edges:
            if _dist(es, ee) < MIN_MEMBER_LENGTH:
                continue
            key = (_pt_key(es), _pt_key(ee))
            rkey = (_pt_key(ee), _pt_key(es))
            if key in seen or rkey in seen:
                continue
            seen.add(key)
            member = self._make_roof_border_member(
                es,
                ee,
                plane,
                roof_info,
                "FASCIA",
                edge_role,
                edge_depth,
            )
            if member is not None:
                members.append(member)
        return members

    # ------------------------------------------------------------------
    #  Ledger -- high-side beam on shed roofs
    # ------------------------------------------------------------------

    def _make_ledger(self, ledger_edges, roof_info):
        """Ledger / header beam at the high side of a shed roof."""
        members = []
        seen = set()
        for ls, le, plane, edge_role, edge_depth in ledger_edges:
            if _dist(ls, le) < MIN_MEMBER_LENGTH:
                continue
            key = (_pt_key(ls), _pt_key(le))
            rkey = (_pt_key(le), _pt_key(ls))
            if key in seen or rkey in seen:
                continue
            seen.add(key)
            member = self._make_roof_border_member(
                ls,
                le,
                plane,
                roof_info,
                "LEDGER",
                edge_role,
                edge_depth,
            )
            if member is not None:
                members.append(member)
        return members

    def _make_roof_border_member(self, start_point, end_point, plane, roof_info, member_type, edge_role, edge_depth):
        """Create a roof border member aligned to the host roof face."""
        from Autodesk.Revit.DB import XYZ

        member_dir = _normalize(end_point - start_point)
        if member_dir is None:
            return None

        family_name = (
            self.config.header_family_name or self.config.stud_family_name)
        type_name = (
            self.config.header_type_name or self.config.stud_type_name)
        member_width, member_depth = self._resolve_roof_member_size(
            family_name,
            type_name,
        )
        layer_top_depth = self._resolve_roof_layer_top_depth(plane)
        roof_normal = getattr(plane, "normal", None)
        roof_normal = _normalize(roof_normal) if roof_normal is not None else None

        rotation = 0.0
        offset = None

        extra_depth = max(0.0, layer_top_depth - max(0.0, edge_depth))
        if roof_normal is not None and extra_depth > 0.0:
            offset = roof_normal.Multiply(-extra_depth)

        if edge_role in ("eave", "ledger"):
            plumb_down = _normalize(
                _project_perpendicular(XYZ.BasisZ.Multiply(-1.0), member_dir)
            )
            outward = getattr(plane, "y_axis", None)
            outward = _normalize(_project_perpendicular(outward, member_dir))
            if outward is not None and edge_role == "ledger":
                outward = outward.Multiply(-1.0)

            if plumb_down is not None and member_depth > 0.0:
                drop_shift = plumb_down.Multiply(member_depth / 2.0)
                offset = drop_shift if offset is None else offset + drop_shift
            if outward is not None and member_width > 0.0:
                outward_shift = outward.Multiply(-member_width / 2.0)
                offset = outward_shift if offset is None else offset + outward_shift
        else:
            desired_up = roof_normal
            reference_up = _beam_reference_up(member_dir)
            if reference_up is not None and desired_up is not None:
                rotation = _signed_angle_about(member_dir, reference_up, desired_up)
            if member_depth > 0.0 and desired_up is not None:
                depth_shift = desired_up.Multiply(-member_depth / 2.0)
                offset = depth_shift if offset is None else offset + depth_shift

            side_axis = _normalize(
                _project_perpendicular(getattr(plane, "x_axis", None), member_dir)
            )
            if side_axis is not None and member_width > 0.0:
                probe_start = start_point if offset is None else start_point + offset
                probe_end = end_point if offset is None else end_point + offset
                best_shifted = self._choose_best_side_shift(
                    roof_info.element,
                    probe_start,
                    probe_end,
                    side_axis,
                    member_width / 2.0,
                )
                if best_shifted is not None:
                    start_point, end_point = best_shifted
                    offset = None

        if offset is not None:
            start_point = start_point + offset
            end_point = end_point + offset

        clipped = self._clip_member_axis_to_roof(roof_info.element, start_point, end_point)
        if clipped is not None:
            start_point, end_point = clipped
        if _dist(start_point, end_point) < MIN_MEMBER_LENGTH:
            return None

        member = FramingMember(FramingMember.HEADER, start_point, end_point)
        member.member_type = member_type
        member.family_name = family_name
        member.type_name = type_name
        member.rotation = rotation
        member.host_kind = roof_info.kind
        member.host_id = roof_info.element_id
        return member

    @staticmethod
    def _resolve_roof_layer_top_depth(plane):
        """Return the depth from roof exterior to the top of the target layer."""
        target_layer = getattr(plane, "target_layer", None)
        if target_layer is None:
            return 0.0
        try:
            return max(0.0, float(getattr(target_layer, "start_depth", 0.0)))
        except Exception:
            return 0.0

    def _resolve_roof_member_size(self, family_name, type_name):
        """Resolve member thickness and depth from the type or nominal size."""
        depth = self.get_type_depth(family_name, type_name)
        width = None

        text = "{0} {1}".format(family_name or "", type_name or "").lower()
        for nominal, dimensions in LUMBER_ACTUAL.items():
            if nominal.lower() in text:
                width = inches_to_feet(dimensions[0])
                if depth is None or depth <= 0.0:
                    depth = inches_to_feet(dimensions[1])
                break

        if depth is not None and depth > 0.0:
            if width is None or width <= 0.0:
                width = PLATE_THICKNESS
            return width, depth

        match = re.search(r"\b2x(2|3|4|6|8|10|12)\b", text)
        if match:
            nominal = "2x{0}".format(match.group(1))
            dims = LUMBER_ACTUAL.get(nominal)
            if dims is not None:
                return inches_to_feet(dims[0]), inches_to_feet(dims[1])

        return PLATE_THICKNESS, 0.0

    # ------------------------------------------------------------------
    #  Truss placement
    # ------------------------------------------------------------------

    def _calc_truss_positions(self, roof_info):
        """Place trusses at OC spacing perpendicular to ridge, eave-to-eave."""
        members = []
        planes = roof_info.planes
        spacing = self.config.stud_spacing_ft
        if spacing <= 0:
            return members

        try:
            ridge_edges = _find_ridge_edges(planes)
        except Exception:
            ridge_edges = []

        sloped = [p for p in planes if p.normal.Z < FLAT_THRESHOLD]

        if not ridge_edges:
            if sloped:
                return self._make_rafters_scanline(
                    sloped[0], spacing, roof_info)
            return members

        seen = set()
        for rs, re, pa, pb in ridge_edges:
            ridge_dir = re - rs
            ridge_len = ridge_dir.GetLength()
            if ridge_len < MIN_MEMBER_LENGTH:
                continue
            ridge_unit = ridge_dir.Multiply(1.0 / ridge_len)

            class_a = _classify_boundary_edges(pa, ridge_edges, (rs, re))
            class_b = _classify_boundary_edges(pb, ridge_edges, (rs, re))
            eave_a = class_a["eave"]
            eave_b = class_b["eave"]
            if not eave_a or not eave_b:
                continue

            d = 0.0
            while d <= ridge_len + 1e-9:
                ridge_pt = rs + ridge_unit.Multiply(min(d, ridge_len))
                foot_a = _project_to_best_eave(ridge_pt, eave_a, rs, re)
                foot_b = _project_to_best_eave(ridge_pt, eave_b, rs, re)
                if foot_a is None or foot_b is None:
                    d += spacing
                    continue

                if _dist(foot_a, foot_b) >= MIN_MEMBER_LENGTH:
                    key = (_pt_key(foot_a), _pt_key(foot_b))
                    rkey = (_pt_key(foot_b), _pt_key(foot_a))
                    if key not in seen and rkey not in seen:
                        seen.add(key)
                        m = FramingMember(FramingMember.STUD, foot_a, foot_b)
                        m.member_type = "TRUSS"
                        m.family_name = self.config.stud_family_name
                        m.type_name = self.config.stud_type_name
                        m.rotation = 0.0
                        m.host_kind = roof_info.kind
                        m.host_id = roof_info.element_id
                        members.append(m)

                d += spacing

        return members