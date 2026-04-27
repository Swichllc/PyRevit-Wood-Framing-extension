# -*- coding: utf-8 -*-
"""Clean multi-slope roof framing planner and V2 placement engine.

This module intentionally does not reuse the legacy placement logic in
`wf_roof.py`. Its job is to produce a reliable per-bay / per-field plan for
multi-slope roofs so the eventual placement step can use Revit-native beam
system workflows instead of direct member patching.
"""

import math
import re

from wf_config import LUMBER_ACTUAL
from wf_geometry import FramingMember, inches_to_feet
from wf_host import analyze_roof_host, _scanline_intervals
from wf_placement import BaseFramingEngine
from wf_schedule_utils import apply_bom_metadata
from wf_tracking import tag_instance


FLAT_THRESHOLD = 0.9998
EDGE_TOL = 1.0 / 12.0
PROFILE_MATCH_TOL = 0.125 / 12.0
PARALLEL_DOT_THRESHOLD = 0.5
MIN_SEGMENT_LENGTH = 1.0 / 12.0
MIN_MEMBER_LENGTH = inches_to_feet(1.0)
RAFTER_TIE_FRACTION = 1.0 / 3.0
COLLAR_TIE_FRACTION = 1.0 / 3.0
MAX_COLLAR_TIE_SPACING = inches_to_feet(48.0)
KICKER_FRACTION = 0.25
V2_BUILD_TAG = "2026-04-22-tie-full-rewrite-1"


class RoofBayPlan(object):
    def __init__(self, index, ridge_start, ridge_end, plane_a_index, plane_b_index):
        self.index = index
        self.ridge_start = ridge_start
        self.ridge_end = ridge_end
        self.plane_a_index = plane_a_index
        self.plane_b_index = plane_b_index
        self.notes = []


class RoofFieldPlan(object):
    def __init__(self, index, ridge_index, plane_index, side_label):
        self.index = index
        self.ridge_index = ridge_index
        self.plane_index = plane_index
        self.side_label = side_label
        self.system_mode = "non_planar_sketch_beam_system"
        self.layout_rule = "fixed_distance"
        self.justification = "direction_line"
        self.eave_edges = []
        self.rake_edges = []
        self.ledger_edges = []
        self.direction_start = None
        self.direction_end = None
        self.split_required = False
        self.notes = []


class RoofPlanV2(object):
    def __init__(self, roof_info, roof_type, sloped_plane_count, ridge_count):
        self.roof_info = roof_info
        self.roof_type = roof_type
        self.sloped_plane_count = sloped_plane_count
        self.ridge_count = ridge_count
        self.supported = False
        self.plan_mode = "analysis_only"
        self.bays = []
        self.fields = []
        self.warnings = []
        self.recommendations = []


def _pt_key(point, decimals=4):
    return (
        round(point.X, decimals),
        round(point.Y, decimals),
        round(point.Z, decimals),
    )


def _dist(first, second):
    return first.DistanceTo(second)


def _normalize(vector):
    length = vector.GetLength()
    if length < 1e-9:
        return None
    return vector.Multiply(1.0 / length)


def _midpoint(first, second):
    return first + (second - first).Multiply(0.5)


def _lerp(first, second, t):
    return first + (second - first).Multiply(t)


def _segment_length(start_point, end_point):
    try:
        return (end_point - start_point).GetLength()
    except Exception:
        return 0.0


def _project_perpendicular(vector, axis):
    axis_unit = _normalize(axis)
    if axis_unit is None:
        return None
    return vector - axis_unit.Multiply(vector.DotProduct(axis_unit))


def _beam_reference_up(member_dir):
    from Autodesk.Revit.DB import XYZ

    reference_up = _normalize(_project_perpendicular(XYZ.BasisZ, member_dir))
    if reference_up is not None:
        return reference_up
    return _normalize(_project_perpendicular(XYZ.BasisX, member_dir))


def _signed_angle_about(axis, start_vec, end_vec):
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
    if desired_up is None:
        return 0.0

    reference_up = _beam_reference_up(member_dir)
    desired_up = _normalize(_project_perpendicular(desired_up, member_dir))
    if reference_up is None or desired_up is None:
        return 0.0
    return _signed_angle_about(member_dir, reference_up, desired_up)


def _offset_from_surface(point, normal, depth):
    return point + normal.Multiply(-depth)


def _resolve_roof_layer_top_depth(plane):
    target_layer = getattr(plane, "target_layer", None)
    if target_layer is None:
        return 0.0
    try:
        return max(0.0, float(getattr(target_layer, "start_depth", 0.0)))
    except Exception:
        return 0.0


def _set_beam_system_elevation(beam_system, offset):
    try:
        beam_system.Elevation = offset
        return True
    except Exception:
        pass

    try:
        from Autodesk.Revit.DB import BuiltInParameter

        parameter = beam_system.get_Parameter(BuiltInParameter.INSTANCE_ELEVATION_PARAM)
        if parameter is not None and not parameter.IsReadOnly:
            parameter.Set(offset)
            return True
    except Exception:
        pass

    for name in (
        "Elevation",
        "Elevation from Level",
        "Offset From Level",
        "Height Offset From Level",
    ):
        try:
            parameter = beam_system.LookupParameter(name)
            if parameter is not None and not parameter.IsReadOnly:
                parameter.Set(offset)
                return True
        except Exception:
            pass

    return False


def _stations_along_segment(start_point, end_point, spacing):
    segment_len = _segment_length(start_point, end_point)
    if segment_len < MIN_MEMBER_LENGTH or spacing <= 0.0:
        return []

    direction = _normalize(end_point - start_point)
    if direction is None:
        return []

    stations = [0.0]
    distance = spacing
    while distance < segment_len - 1e-9:
        stations.append(distance)
        distance += spacing
    if abs(stations[-1] - segment_len) > 1e-6:
        stations.append(segment_len)

    return [start_point + direction.Multiply(distance) for distance in stations]


def _project_to_edge(point, edge_start, edge_end):
    edge_dir = edge_end - edge_start
    edge_len = edge_dir.GetLength()
    if edge_len < 1e-9:
        return edge_start, False

    edge_unit = edge_dir.Multiply(1.0 / edge_len)
    raw_t = (point - edge_start).DotProduct(edge_unit)
    clamped_t = max(0.0, min(edge_len, raw_t))
    projected = edge_start + edge_unit.Multiply(clamped_t)
    is_interior = abs(clamped_t - raw_t) <= 1e-9
    return projected, is_interior


def _ridge_station_on_segment(point, ridge_start, ridge_end):
    ridge_dir = ridge_end - ridge_start
    ridge_len = ridge_dir.GetLength()
    if ridge_len < 1e-9:
        return 0.0
    ridge_unit = ridge_dir.Multiply(1.0 / ridge_len)
    return (point - ridge_start).DotProduct(ridge_unit)


def _segment_covers_ridge_station(edge_start, edge_end, ridge_start, ridge_end, ridge_station):
    edge_station_start = _ridge_station_on_segment(edge_start, ridge_start, ridge_end)
    edge_station_end = _ridge_station_on_segment(edge_end, ridge_start, ridge_end)
    station_min = min(edge_station_start, edge_station_end)
    station_max = max(edge_station_start, edge_station_end)
    return (station_min - EDGE_TOL) <= ridge_station <= (station_max + EDGE_TOL)


def _project_to_best_support(ridge_point, support_edges, ridge_start, ridge_end):
    if not support_edges:
        return None

    ridge_station = _ridge_station_on_segment(ridge_point, ridge_start, ridge_end)
    best_point = None
    best_covers_station = False
    best_interior = False
    best_distance = None

    for edge_start, edge_end in support_edges:
        projected, is_interior = _project_to_edge(ridge_point, edge_start, edge_end)
        covers_station = _segment_covers_ridge_station(
            edge_start,
            edge_end,
            ridge_start,
            ridge_end,
            ridge_station,
        )
        distance = _dist(ridge_point, projected)

        if best_point is None:
            best_point = projected
            best_covers_station = covers_station
            best_interior = is_interior
            best_distance = distance
            continue

        if covers_station and not best_covers_station:
            best_point = projected
            best_covers_station = True
            best_interior = is_interior
            best_distance = distance
            continue

        if is_interior and not best_interior:
            best_point = projected
            best_covers_station = covers_station
            best_interior = True
            best_distance = distance
            continue

        if (covers_station == best_covers_station
                and is_interior == best_interior
                and distance < best_distance):
            best_point = projected
            best_covers_station = covers_station
            best_interior = is_interior
            best_distance = distance

    return best_point


def _cross_2d(first, second):
    return (first[0] * second[1]) - (first[1] * second[0])


def _support_point_along_local_axis(ridge_point, support_edges, plane, axis_u):
    if ridge_point is None or not support_edges or plane is None or axis_u is None:
        return None

    ridge_local = _to_plane_local(ridge_point, plane)
    ray_dir = (-axis_u[0], -axis_u[1])
    ray_len = math.sqrt((ray_dir[0] * ray_dir[0]) + (ray_dir[1] * ray_dir[1]))
    if ray_len < 1e-9:
        return None

    best_hit = None
    best_t = None
    for edge_start, edge_end in support_edges:
        start_local = _to_plane_local(edge_start, plane)
        end_local = _to_plane_local(edge_end, plane)
        edge_dir = (
            end_local[0] - start_local[0],
            end_local[1] - start_local[1],
        )
        denom = _cross_2d(ray_dir, edge_dir)
        if abs(denom) <= 1e-9:
            continue

        diff = (
            start_local[0] - ridge_local[0],
            start_local[1] - ridge_local[1],
        )
        ray_t = _cross_2d(diff, edge_dir) / denom
        edge_t = _cross_2d(diff, ray_dir) / denom
        if ray_t < -1e-9:
            continue
        if edge_t < -1e-9 or edge_t > 1.0 + 1e-9:
            continue

        if best_t is None or ray_t < best_t:
            best_t = ray_t
            best_hit = (
                ridge_local[0] + (ray_dir[0] * ray_t),
                ridge_local[1] + (ray_dir[1] * ray_t),
            )

    if best_hit is None:
        return None

    return _surface_point(plane, best_hit[0], best_hit[1])


def _trim_segment_ends(start_point, end_point, trim_each_end):
    if trim_each_end <= 0.0:
        return start_point, end_point

    direction = _normalize(end_point - start_point)
    if direction is None:
        return start_point, end_point

    length = _dist(start_point, end_point)
    if length <= (trim_each_end * 2.0) + MIN_MEMBER_LENGTH:
        return start_point, end_point

    return (
        start_point + direction.Multiply(trim_each_end),
        end_point - direction.Multiply(trim_each_end),
    )


def _match_by_ridge_station(lines_a, lines_b, max_station_delta=None):
    """Greedily pair rafter lines from opposite sides of a bay by ridge station.

    Each line is a ``(ridge_station, eave_point, ridge_point)`` tuple.
    Returns matched pairs ``[((s_a, ea, ra), (s_b, eb, rb)), ...]``.
    Each rafter on side B is used at most once, preventing duplicate ties when
    the rafter counts are unequal.  When ``max_station_delta`` is provided,
    pairs whose station offset exceeds the tolerance are skipped instead of
    force-matched.
    """
    if not lines_a or not lines_b:
        return []
    remaining_b = list(lines_b)
    matched = []
    for line_a in lines_a:
        if not remaining_b:
            break
        station_a = line_a[0]
        best_index = min(
            range(len(remaining_b)),
            key=lambda i: abs(remaining_b[i][0] - station_a),
        )
        if max_station_delta is not None:
            if abs(remaining_b[best_index][0] - station_a) > max_station_delta:
                continue
        matched.append((line_a, remaining_b[best_index]))
        remaining_b.pop(best_index)
    return matched


def _match_lines_by_station(lines_a, lines_b, tolerance):
    """Pair two sorted station lists in-order within a strict tolerance."""
    if not lines_a or not lines_b:
        return []

    tol = max(0.0, float(tolerance or 0.0))
    pairs = []
    index_a = 0
    index_b = 0
    while index_a < len(lines_a) and index_b < len(lines_b):
        station_a = lines_a[index_a][0]
        station_b = lines_b[index_b][0]
        delta = station_a - station_b
        if abs(delta) <= tol:
            pairs.append((lines_a[index_a], lines_b[index_b]))
            index_a += 1
            index_b += 1
        elif delta < 0.0:
            index_a += 1
        else:
            index_b += 1
    return pairs


def _point_on_segment_at_z(start_point, end_point, target_z):
    """Return the point on a segment at target Z, or None when out of range."""
    delta_z = end_point.Z - start_point.Z
    if abs(delta_z) <= 1e-9:
        if abs(target_z - start_point.Z) <= EDGE_TOL:
            return _midpoint(start_point, end_point)
        return None

    t = (target_z - start_point.Z) / delta_z
    if t < -1e-6 or t > 1.0 + 1e-6:
        return None
    if t < 0.0:
        t = 0.0
    elif t > 1.0:
        t = 1.0
    return _lerp(start_point, end_point, t)


def _common_z_interval(seg_a_start, seg_a_end, seg_b_start, seg_b_end):
    """Return the overlapping Z range between two segments."""
    min_a = min(seg_a_start.Z, seg_a_end.Z)
    max_a = max(seg_a_start.Z, seg_a_end.Z)
    min_b = min(seg_b_start.Z, seg_b_end.Z)
    max_b = max(seg_b_start.Z, seg_b_end.Z)
    return max(min_a, min_b), min(max_a, max_b)


def _member_depth_from_text(text):
    if not text:
        return None

    match = re.search(r"\b2x(2|3|4|6|8|10|12)\b", text)
    if not match:
        return None

    nominal = "2x{0}".format(match.group(1))
    dims = LUMBER_ACTUAL.get(nominal)
    if dims is None:
        return None
    return inches_to_feet(dims[1])


def _member_width_from_text(text):
    if not text:
        return None

    match = re.search(r"\b2x(2|3|4|6|8|10|12)\b", text)
    if not match:
        return None

    nominal = "2x{0}".format(match.group(1))
    dims = LUMBER_ACTUAL.get(nominal)
    if dims is None:
        return None
    return inches_to_feet(dims[0])


def _surface_point(plane, local_x, local_y):
    return (
        plane.origin
        + plane.x_axis.Multiply(local_x)
        + plane.y_axis.Multiply(local_y)
    )


def _sloped_planes(planes):
    return [plane for plane in planes if plane.normal.Z < FLAT_THRESHOLD]


def _classify_roof_type(planes):
    sloped = _sloped_planes(planes)
    flat_planes = [plane for plane in planes if plane.normal.Z >= FLAT_THRESHOLD]
    if not sloped:
        return "flat"
    if len(sloped) == 1:
        return "shed"
    if len(sloped) == 2 and not flat_planes:
        return "gable"
    if len(sloped) == 4 and not flat_planes:
        return "hip"
    if len(sloped) >= 2 and flat_planes:
        return "dutch"
    return "complex"


def _points_near(first, second, tolerance=PROFILE_MATCH_TOL):
    try:
        return first.DistanceTo(second) <= tolerance
    except Exception:
        return False


def _same_segment(start_a, end_a, start_b, end_b, tolerance=PROFILE_MATCH_TOL):
    return (
        (_points_near(start_a, start_b, tolerance)
         and _points_near(end_a, end_b, tolerance))
        or (_points_near(start_a, end_b, tolerance)
            and _points_near(end_a, start_b, tolerance))
    )


def _find_ridge_segments(planes):
    sloped = _sloped_planes(planes)
    if len(sloped) < 2:
        return []

    edges = []
    for plane in sloped:
        for loop in getattr(plane, "boundary_loops_local", []) or []:
            count = len(loop)
            for index in range(count):
                start_local = loop[index]
                end_local = loop[(index + 1) % count]
                start_world = _surface_point(plane, start_local[0], start_local[1])
                end_world = _surface_point(plane, end_local[0], end_local[1])
                edges.append((start_world, end_world, plane))

    ridges = []
    used = set()
    for first_index in range(len(edges)):
        if first_index in used:
            continue
        start_a, end_a, plane_a = edges[first_index]
        if _segment_length(start_a, end_a) < MIN_SEGMENT_LENGTH:
            continue
        for second_index in range(first_index + 1, len(edges)):
            if second_index in used:
                continue
            start_b, end_b, plane_b = edges[second_index]
            if plane_a is plane_b:
                continue
            if not _same_segment(start_a, end_a, start_b, end_b):
                continue
            used.add(first_index)
            used.add(second_index)
            ridges.append((start_a, end_a, plane_a, plane_b))
            break

    unique = []
    seen = set()
    for start_point, end_point, plane_a, plane_b in ridges:
        key = (_pt_key(start_point), _pt_key(end_point))
        reverse_key = (_pt_key(end_point), _pt_key(start_point))
        if key in seen or reverse_key in seen:
            continue
        seen.add(key)
        unique.append((start_point, end_point, plane_a, plane_b))

    unique.sort(key=lambda item: _segment_length(item[0], item[1]), reverse=True)
    return unique


def _get_non_ridge_edges(plane, ridge_segments):
    ridge_keys = set()
    for ridge_start, ridge_end, plane_a, plane_b in ridge_segments:
        if plane is not plane_a and plane is not plane_b:
            continue
        ridge_keys.add((_pt_key(ridge_start), _pt_key(ridge_end)))
        ridge_keys.add((_pt_key(ridge_end), _pt_key(ridge_start)))

    edges = []
    for loop in getattr(plane, "boundary_loops_local", []) or []:
        count = len(loop)
        for index in range(count):
            start_local = loop[index]
            end_local = loop[(index + 1) % count]
            start_world = _surface_point(plane, start_local[0], start_local[1])
            end_world = _surface_point(plane, end_local[0], end_local[1])
            if ((_pt_key(start_world), _pt_key(end_world)) in ridge_keys):
                continue
            edges.append((start_world, end_world))
    return edges


def _classify_edges_for_ridge(plane, ridge_segments, ridge_segment):
    non_ridge = _get_non_ridge_edges(plane, ridge_segments)
    if not non_ridge:
        return {"eave": [], "rake": [], "ledger": []}

    ridge_start, ridge_end = ridge_segment
    ridge_dir = _normalize(ridge_end - ridge_start)
    if ridge_dir is None:
        return {"eave": [], "rake": [], "ledger": []}

    parallel = []
    perpendicular = []
    for start_point, end_point in non_ridge:
        edge_dir = _normalize(end_point - start_point)
        if edge_dir is None:
            continue
        avg_z = (start_point.Z + end_point.Z) / 2.0
        dot = abs(edge_dir.DotProduct(ridge_dir))
        if dot >= PARALLEL_DOT_THRESHOLD:
            parallel.append((start_point, end_point, avg_z))
        else:
            perpendicular.append((start_point, end_point))

    eave = []
    ledger = []
    if parallel:
        z_values = [avg_z for _, _, avg_z in parallel]
        z_min = min(z_values)
        z_max = max(z_values)
        z_mid = (z_min + z_max) / 2.0
        if z_max - z_min < EDGE_TOL:
            eave = [(start_point, end_point) for start_point, end_point, _ in parallel]
        else:
            for start_point, end_point, avg_z in parallel:
                if avg_z <= z_mid:
                    eave.append((start_point, end_point))
                else:
                    ledger.append((start_point, end_point))

    return {
        "eave": eave,
        "rake": perpendicular,
        "ledger": ledger,
    }


def _edge_midpoint(edge):
    return _midpoint(edge[0], edge[1])


def _longest_edge(edges):
    if not edges:
        return None
    return sorted(edges, key=lambda edge: _segment_length(edge[0], edge[1]), reverse=True)[0]


def _direction_line_from_eave_to_ridge(eave_edges, ridge_start, ridge_end):
    main_eave = _longest_edge(eave_edges)
    if main_eave is None:
        return None, None
    eave_mid = _edge_midpoint(main_eave)
    ridge_mid = _midpoint(ridge_start, ridge_end)
    if _segment_length(eave_mid, ridge_mid) < MIN_SEGMENT_LENGTH:
        return None, None
    return eave_mid, ridge_mid


def _normalize_2d(vector):
    length = math.sqrt((vector[0] * vector[0]) + (vector[1] * vector[1]))
    if length < 1e-9:
        return None
    return (vector[0] / length, vector[1] / length)


def _dot_2d(first, second):
    return (first[0] * second[0]) + (first[1] * second[1])


def _dist_2d(first, second):
    dx = first[0] - second[0]
    dy = first[1] - second[1]
    return math.sqrt((dx * dx) + (dy * dy))


def _points_near_2d(first, second, tolerance=PROFILE_MATCH_TOL):
    return _dist_2d(first, second) <= tolerance


def _same_segment_2d(start_a, end_a, start_b, end_b, tolerance=PROFILE_MATCH_TOL):
    return (
        (_points_near_2d(start_a, start_b, tolerance)
         and _points_near_2d(end_a, end_b, tolerance))
        or (_points_near_2d(start_a, end_b, tolerance)
            and _points_near_2d(end_a, start_b, tolerance))
    )


def _to_plane_local(point, plane):
    vector = point - plane.origin
    return (
        vector.DotProduct(plane.x_axis),
        vector.DotProduct(plane.y_axis),
    )


def _field_loop_local_for_ridge(plane, ridge_start, ridge_end):
    ridge_start_local = _to_plane_local(ridge_start, plane)
    ridge_end_local = _to_plane_local(ridge_end, plane)

    for loop in getattr(plane, "boundary_loops_local", []) or []:
        count = len(loop)
        for index in range(count):
            seg_start = loop[index]
            seg_end = loop[(index + 1) % count]
            if not _same_segment_2d(seg_start, seg_end, ridge_start_local, ridge_end_local):
                continue

            if (_points_near_2d(seg_start, ridge_start_local)
                    and _points_near_2d(seg_end, ridge_end_local)):
                chain = []
                walk_index = (index + 2) % count
                while not _points_near_2d(loop[walk_index], ridge_start_local):
                    chain.append(loop[walk_index])
                    walk_index = (walk_index + 1) % count
                return [ridge_start_local, ridge_end_local] + chain

            chain = []
            walk_index = (index - 1) % count
            while not _points_near_2d(loop[walk_index], ridge_start_local):
                chain.append(loop[walk_index])
                walk_index = (walk_index - 1) % count
            return [ridge_start_local, ridge_end_local] + chain

    return None


def _field_loop_local_from_support_edge(field, plane, ridge_start, ridge_end):
    support_edge = _longest_edge(
        list(getattr(field, "eave_edges", []) or getattr(field, "ledger_edges", []) or [])
    )
    if support_edge is None:
        return None

    ridge_start_local = _to_plane_local(ridge_start, plane)
    ridge_end_local = _to_plane_local(ridge_end, plane)
    support_start_local = _to_plane_local(support_edge[0], plane)
    support_end_local = _to_plane_local(support_edge[1], plane)

    direct_score = (
        _dist_2d(ridge_start_local, support_start_local)
        + _dist_2d(ridge_end_local, support_end_local)
    )
    flipped_score = (
        _dist_2d(ridge_start_local, support_end_local)
        + _dist_2d(ridge_end_local, support_start_local)
    )
    if flipped_score < direct_score:
        support_start_local, support_end_local = support_end_local, support_start_local

    return [
        ridge_start_local,
        ridge_end_local,
        support_end_local,
        support_start_local,
    ]


def _resolved_field_loop_local(field, plane, ridge_start, ridge_end):
    field_loop_local = getattr(field, "loop_local_override", None)
    if not field_loop_local:
        field_loop_local = _field_loop_local_for_ridge(plane, ridge_start, ridge_end)
    if not field_loop_local:
        field_loop_local = _field_loop_local_from_support_edge(field, plane, ridge_start, ridge_end)
    return field_loop_local


def _transform_loop_2d(loop_local, axis_u, axis_v):
    transformed = []
    for point in loop_local:
        transformed.append((
            _dot_2d(point, axis_u),
            _dot_2d(point, axis_v),
        ))
    return transformed


def _inverse_transform_2d(coord_u, coord_v, axis_u, axis_v):
    return (
        (axis_u[0] * coord_u) + (axis_v[0] * coord_v),
        (axis_u[1] * coord_u) + (axis_v[1] * coord_v),
    )


def _field_direction_axes(field, plane, ridge_start, ridge_end):
    ridge_start_local = _to_plane_local(ridge_start, plane)
    ridge_end_local = _to_plane_local(ridge_end, plane)
    ridge_mid_local = (
        (ridge_start_local[0] + ridge_end_local[0]) / 2.0,
        (ridge_start_local[1] + ridge_end_local[1]) / 2.0,
    )

    main_eave = _longest_edge(getattr(field, "eave_edges", []) or [])
    if main_eave is not None:
        eave_start_local = _to_plane_local(main_eave[0], plane)
        eave_end_local = _to_plane_local(main_eave[1], plane)
        eave_dir = _normalize_2d((
            eave_end_local[0] - eave_start_local[0],
            eave_end_local[1] - eave_start_local[1],
        ))
        if eave_dir is not None:
            axis_u = (-eave_dir[1], eave_dir[0])
        else:
            axis_u = None
    else:
        axis_u = None

    direction_start = getattr(field, "direction_start", None)
    if axis_u is None:
        ridge_dir = _normalize_2d((
            ridge_end_local[0] - ridge_start_local[0],
            ridge_end_local[1] - ridge_start_local[1],
        ))
        if ridge_dir is None:
            return None, None
        axis_u = (-ridge_dir[1], ridge_dir[0])

    support_local = None
    if main_eave is not None:
        support_local = _to_plane_local(_edge_midpoint(main_eave), plane)
    elif direction_start is not None:
        support_local = _to_plane_local(direction_start, plane)
    if support_local is not None:
        ridge_u = _dot_2d(ridge_mid_local, axis_u)
        support_u = _dot_2d(support_local, axis_u)
        if support_u > ridge_u:
            axis_u = (-axis_u[0], -axis_u[1])

    axis_v = (-axis_u[1], axis_u[0])
    return axis_u, axis_v


def _field_scan_stations(loop_uv, spacing):
    min_v = min(point[1] for point in loop_uv)
    max_v = max(point[1] for point in loop_uv)
    if max_v - min_v < MIN_MEMBER_LENGTH:
        return []

    stations = []
    coord = min_v
    while coord <= max_v + 1e-9:
        stations.append(coord)
        coord += spacing
    return stations


def _segment_length_between_points(start_point, end_point):
    if start_point is None or end_point is None:
        return 0.0
    return _segment_length(start_point, end_point)


def _segment_level_delta(start_point, end_point):
    if start_point is None or end_point is None:
        return float("inf")
    return abs(float(start_point.Z) - float(end_point.Z))


def _segment_average_z(start_point, end_point):
    if start_point is None or end_point is None:
        return float("-inf")
    return (float(start_point.Z) + float(end_point.Z)) / 2.0


def _longest_edge_length(edges):
    if not edges:
        return 0.0
    return max(_segment_length(edge[0], edge[1]) for edge in edges)


def _field_selection_score(plan, field):
    bay = None
    if 0 <= field.ridge_index < len(getattr(plan, "bays", []) or []):
        bay = plan.bays[field.ridge_index]

    support_start = getattr(bay, "ridge_start", None)
    support_end = getattr(bay, "ridge_end", None)
    level_delta = _segment_level_delta(support_start, support_end)
    average_z = _segment_average_z(support_start, support_end)
    direction_length = _segment_length_between_points(
        getattr(field, "direction_start", None),
        getattr(field, "direction_end", None),
    )
    eave_length = _longest_edge_length(getattr(field, "eave_edges", []) or [])
    return (
        1 if level_delta <= EDGE_TOL else 0,
        average_z,
        -level_delta,
        eave_length,
        direction_length,
    )


def _placement_fields_for_plan(plan):
    grouped = {}
    ordered_plane_indexes = []
    for field in plan.fields:
        plane_index = field.plane_index
        if plane_index not in grouped:
            grouped[plane_index] = []
            ordered_plane_indexes.append(plane_index)
        grouped[plane_index].append(field)

    placement_fields = []
    sloped = _sloped_planes(getattr(plan.roof_info, "planes", []) or [])
    for plane_index in ordered_plane_indexes:
        plane_fields = grouped[plane_index]
        if len(plane_fields) > 1:
            collapsed = _collapsed_field_for_plane(plan, sloped, plane_index, plane_fields)
            if collapsed is not None:
                placement_fields.append(collapsed)
            continue
        placement_fields.extend(plane_fields)

    return placement_fields


def _collapsed_field_for_plane(plan, sloped_planes, plane_index, plane_fields):
    if plane_index < 0 or plane_index >= len(sloped_planes):
        return None

    plane = sloped_planes[plane_index]
    best_field = sorted(plane_fields, key=lambda field: _field_selection_score(plan, field), reverse=True)[0]

    collapsed = RoofFieldPlan(best_field.index, best_field.ridge_index, plane_index, "collapsed")
    collapsed.system_mode = best_field.system_mode
    collapsed.layout_rule = best_field.layout_rule
    collapsed.justification = best_field.justification
    collapsed.direction_start = best_field.direction_start
    collapsed.direction_end = best_field.direction_end
    collapsed.eave_edges = list(getattr(best_field, "eave_edges", []) or [])
    collapsed.rake_edges = []
    collapsed.ledger_edges = []
    collapsed.split_required = False
    collapsed.notes = list(getattr(best_field, "notes", []) or [])
    collapsed.notes.append(
        "Placement fallback: selected one primary framing support for this plane to avoid overlapping directions."
    )

    best_loop = getattr(best_field, "loop_local_override", None)
    if not best_loop and 0 <= best_field.ridge_index < len(getattr(plan, "bays", []) or []):
        bay = plan.bays[best_field.ridge_index]
        best_loop = _field_loop_local_for_ridge(
            plane,
            bay.ridge_start,
            bay.ridge_end,
        )

    if best_loop:
        collapsed.loop_local_override = list(best_loop)

    return collapsed


class RoofFramingPlannerV2(object):
    def __init__(self, doc, config=None):
        self.doc = doc
        self.config = config

    def plan_roof(self, roof):
        roof_info = analyze_roof_host(self.doc, roof, self.config)
        if roof_info is None:
            return None
        return self.plan_roof_info(roof_info)

    def plan_roof_info(self, roof_info):
        planes = getattr(roof_info, "planes", []) or []
        sloped = _sloped_planes(planes)
        roof_type = _classify_roof_type(planes)
        ridges = _find_ridge_segments(planes)

        plan = RoofPlanV2(
            roof_info,
            roof_type,
            len(sloped),
            len(ridges),
        )
        plan.recommendations.append(
            "Treat each framing field as a separate non-planar sketched beam system."
        )
        plan.recommendations.append(
            "Do not span multiple dissimilar roof regions with one beam system."
        )
        plan.recommendations.append(
            "Handle ridge boards, ties, fascia, and openings after the beam-system fields are stable."
        )

        if len(sloped) <= 1:
            plan.warnings.append(
                "This roof is not a multi-slope candidate. Use Single-Slope Roof Framing for shed roofs."
            )
            return plan

        if not ridges:
            plan.warnings.append(
                "No ridge segments were detected. Start by validating roof host analysis before any placement work."
            )
            return plan

        plan.supported = True
        plane_index_by_id = {}
        for plane_index, plane in enumerate(sloped):
            plane_index_by_id[id(plane)] = plane_index

        plane_ridge_counts = {}
        for _, _, plane_a, plane_b in ridges:
            plane_ridge_counts[id(plane_a)] = plane_ridge_counts.get(id(plane_a), 0) + 1
            plane_ridge_counts[id(plane_b)] = plane_ridge_counts.get(id(plane_b), 0) + 1

        warning_set = set()
        field_index = 0
        for ridge_index, ridge in enumerate(ridges):
            ridge_start, ridge_end, plane_a, plane_b = ridge
            plane_a_index = plane_index_by_id.get(id(plane_a), -1)
            plane_b_index = plane_index_by_id.get(id(plane_b), -1)

            bay = RoofBayPlan(
                ridge_index,
                ridge_start,
                ridge_end,
                plane_a_index,
                plane_b_index,
            )
            bay.notes.append(
                "Use two separate framing fields here, one on each side of the ridge."
            )
            plan.bays.append(bay)

            for plane, plane_index, side_label in (
                (plane_a, plane_a_index, "A"),
                (plane_b, plane_b_index, "B"),
            ):
                field = RoofFieldPlan(field_index, ridge_index, plane_index, side_label)
                classified = _classify_edges_for_ridge(
                    plane,
                    ridges,
                    (ridge_start, ridge_end),
                )
                field.eave_edges = classified["eave"]
                field.rake_edges = classified["rake"]
                field.ledger_edges = classified["ledger"]
                field.direction_start, field.direction_end = _direction_line_from_eave_to_ridge(
                    field.eave_edges,
                    ridge_start,
                    ridge_end,
                )

                ridge_touch_count = plane_ridge_counts.get(id(plane), 0)
                if ridge_touch_count > 1:
                    field.split_required = True
                    field.notes.append(
                        "This plane touches multiple framing supports; V2 currently picks one primary support per plane until true subfield splitting exists."
                    )
                    warning_text = (
                        "Plane {0} touches multiple framing supports. V2 is temporarily selecting one primary support region for this plane.".format(
                            plane_index,
                        )
                    )
                    if warning_text not in warning_set:
                        warning_set.add(warning_text)
                        plan.warnings.append(warning_text)

                if not field.eave_edges:
                    field.notes.append(
                        "No eave edges parallel to this ridge were detected. Recheck support picking before placement."
                    )
                elif len(field.eave_edges) > 1:
                    field.notes.append(
                        "Multiple eave candidates detected. Prefer sketched supports over automatic boundary generation."
                    )

                if field.direction_start is None or field.direction_end is None:
                    field.notes.append(
                        "Could not resolve a clean direction line from ridge to eave."
                    )

                plan.fields.append(field)
                field_index += 1

        return plan


class RoofFramingEngineV2(BaseFramingEngine):
    """Actual V2 multi-slope placement using the clean field planner."""

    def __init__(self, doc, config):
        BaseFramingEngine.__init__(self, doc, config)
        self.planner = RoofFramingPlannerV2(doc, config)

    def calculate_members(self, roof):
        roof_info = analyze_roof_host(self.doc, roof, self.config)
        if roof_info is None:
            return [], None
        return self.calculate_members_from_roof_info(roof_info)

    def calculate_members_from_roof_info(self, roof_info):
        if roof_info is None:
            return [], None

        plan = self.planner.plan_roof_info(roof_info)
        try:
            roof_info.v2_plan = plan
        except Exception:
            pass

        if plan is None or not plan.supported:
            return [], roof_info

        return self._members_from_plan(plan), roof_info

    def _members_from_plan(self, plan):
        members = []
        members.extend(self._make_ridge_boards_from_plan(plan))
        seen = set()
        for field in _placement_fields_for_plan(plan):
            members.extend(self._make_field_members(plan, field, seen))
        # Requested production scope: rafters, ridge boards, and border members only.
        # Rafter ties are intentionally disabled for now.
        self._record_skipped_secondary_members(plan)
        members.extend(self._make_border_members_from_plan(plan))
        return members

    def _append_plan_warning(self, plan, message):
        warnings = getattr(plan, "warnings", None)
        if warnings is None:
            return
        if message not in warnings:
            warnings.append(message)

    def _record_skipped_secondary_members(self, plan):
        if bool(getattr(self.config, "include_collar_ties", False)):
            self._append_plan_warning(
                plan,
                "V2 collar ties are temporarily disabled; the previous layout was not architecturally correct.",
            )
        if bool(getattr(self.config, "include_ceiling_joists", False)):
            self._append_plan_warning(
                plan,
                "V2 does not place ceiling joists. Use the ceiling framing tool for ceiling members.",
            )
        if bool(getattr(self.config, "include_roof_kickers", False)):
            self._append_plan_warning(
                plan,
                "V2 roof kickers are temporarily disabled; the previous joist-based layout produced incorrect family, count, and length.",
            )

    def _ridge_family_name(self):
        return self.config.header_family_name or self.config.stud_family_name

    def _ridge_type_name(self):
        return self.config.header_type_name or self.config.stud_type_name

    def _edge_family_name(self):
        return (
            getattr(self.config, "roof_edge_family_name", None)
            or self.config.header_family_name
            or self.config.stud_family_name
        )

    def _edge_type_name(self):
        return (
            getattr(self.config, "roof_edge_type_name", None)
            or self.config.header_type_name
            or self.config.stud_type_name
        )

    def _make_ridge_boards_from_plan(self, plan):
        members = []
        seen = set()
        host_kind = getattr(plan.roof_info, "kind", "roof")
        host_id = getattr(plan.roof_info, "element_id", None)
        sloped = _sloped_planes(getattr(plan.roof_info, "planes", []) or [])

        for bay in plan.bays:
            plane = None
            if 0 <= bay.plane_a_index < len(sloped):
                plane = sloped[bay.plane_a_index]
            elif 0 <= bay.plane_b_index < len(sloped):
                plane = sloped[bay.plane_b_index]

            control_depth = self._resolve_board_center_depth(
                plane,
                self._ridge_family_name(),
                self._ridge_type_name(),
            )

            from Autodesk.Revit.DB import XYZ
            ridge_start = XYZ(bay.ridge_start.X, bay.ridge_start.Y, bay.ridge_start.Z - control_depth)
            ridge_end = XYZ(bay.ridge_end.X, bay.ridge_end.Y, bay.ridge_end.Z - control_depth)
            if _dist(ridge_start, ridge_end) < MIN_MEMBER_LENGTH:
                continue
            key = (_pt_key(ridge_start), _pt_key(ridge_end))
            reverse_key = (_pt_key(ridge_end), _pt_key(ridge_start))
            if key in seen or reverse_key in seen:
                continue
            seen.add(key)

            member = FramingMember(FramingMember.HEADER, ridge_start, ridge_end)
            member.member_type = "RIDGE_BOARD_V2"
            member.family_name = self._ridge_family_name()
            member.type_name = self._ridge_type_name()
            member.rotation = 0.0
            member.host_kind = host_kind
            member.host_id = host_id
            members.append(member)
        return members

    def _make_border_members_from_plan(self, plan):
        members = []
        seen = set()
        sloped = _sloped_planes(getattr(plan.roof_info, "planes", []) or [])

        for field in plan.fields:
            if field.plane_index < 0 or field.plane_index >= len(sloped):
                continue
            plane = sloped[field.plane_index]
            members.extend(self._make_border_members_for_edges(
                plane,
                field.eave_edges,
                "FASCIA_V2",
                seen,
            ))
            members.extend(self._make_border_members_for_edges(
                plane,
                field.rake_edges,
                "FASCIA_V2",
                seen,
            ))
            members.extend(self._make_border_members_for_edges(
                plane,
                field.ledger_edges,
                "LEDGER_V2",
                seen,
            ))

        return members

    def _make_border_members_for_edges(self, plane, edges, member_type, seen):
        members = []
        if not edges:
            return members

        control_depth = self._resolve_board_center_depth(
            plane,
            self._edge_family_name(),
            self._edge_type_name(),
        )
        for start_edge, end_edge in edges:
            key = (_pt_key(start_edge), _pt_key(end_edge), member_type)
            reverse_key = (_pt_key(end_edge), _pt_key(start_edge), member_type)
            if key in seen or reverse_key in seen:
                continue
            seen.add(key)

            start_point = _offset_from_surface(start_edge, plane.normal, control_depth)
            end_point = _offset_from_surface(end_edge, plane.normal, control_depth)
            if _dist(start_point, end_point) < MIN_MEMBER_LENGTH:
                continue

            member = FramingMember(FramingMember.HEADER, start_point, end_point)
            member.member_type = member_type
            member.family_name = self._edge_family_name()
            member.type_name = self._edge_type_name()
            member.rotation = 0.0
            member.host_kind = plane.kind
            member.host_id = plane.element_id
            members.append(member)

        return members

    def _make_field_members(self, plan, field, seen):
        spacing = self.config.stud_spacing_ft
        if spacing <= 0.0:
            return []

        sloped = _sloped_planes(getattr(plan.roof_info, "planes", []) or [])
        if field.plane_index < 0 or field.plane_index >= len(sloped):
            return []
        if field.ridge_index < 0 or field.ridge_index >= len(plan.bays):
            return []

        plane = sloped[field.plane_index]
        bay = plan.bays[field.ridge_index]
        ridge_start = bay.ridge_start
        ridge_end = bay.ridge_end
        field_loop_local = _resolved_field_loop_local(field, plane, ridge_start, ridge_end)
        if not field_loop_local or len(field_loop_local) < 3:
            return []

        axis_u, axis_v = _field_direction_axes(field, plane, ridge_start, ridge_end)
        if axis_u is None or axis_v is None:
            return []

        # Instead of creating line segments...
        # Just create one FramingMember acting as a proxy for BeamSystem
        member = FramingMember(FramingMember.STUD, ridge_start, ridge_end) # dummy points
        member.member_type = "BEAM_SYSTEM_V2"
        member.family_name = self.config.stud_family_name
        member.type_name = self.config.stud_type_name
        member.plane = plane
        member.field_loop_local = field_loop_local
        member.direction_local = axis_u
        member.spacing = self.config.stud_spacing_ft
        member.host_kind = plane.kind
        member.host_id = plane.element_id
        
        return [member]

    def _field_for_bay_side(self, plan, bay_index, plane_index):
        for field in getattr(plan, "fields", []) or []:
            if field.ridge_index == bay_index and field.plane_index == plane_index:
                return field
        return None

    def _selected_placement_field_by_plane(self, plan):
        selected = {}
        for field in _placement_fields_for_plan(plan):
            selected[field.plane_index] = field
        return selected

    def _bay_uses_selected_fields(self, plan, bay, selected_fields=None):
        if selected_fields is None:
            selected_fields = self._selected_placement_field_by_plane(plan)

        for plane_index in (bay.plane_a_index, bay.plane_b_index):
            if plane_index < 0:
                continue
            selected = selected_fields.get(plane_index)
            if selected is None:
                return False
            if selected.ridge_index != bay.index:
                return False
        return True

    def _support_edges_for_bay_side(self, plan, bay_index, plane_index):
        field = self._field_for_bay_side(plan, bay_index, plane_index)
        if field is None:
            return []
        return list(
            getattr(field, "eave_edges", [])
            or getattr(field, "ledger_edges", [])
            or []
        )

    def _selected_field_for_bay_side(self, plan, bay, plane_index, selected_fields=None):
        if plane_index < 0:
            return None
        if selected_fields is None:
            selected_fields = self._selected_placement_field_by_plane(plan)
        field = selected_fields.get(plane_index)
        if field is None or field.ridge_index != bay.index:
            return None
        return field

    def _rafter_axis_for_bay_side(self, plan, bay, plane_index):
        sloped = _sloped_planes(getattr(plan.roof_info, "planes", []) or [])
        if plane_index < 0 or plane_index >= len(sloped):
            return None, None

        field = self._field_for_bay_side(plan, bay.index, plane_index)
        if field is None:
            return sloped[plane_index], None

        plane = sloped[plane_index]
        axis_u, _ = _field_direction_axes(field, plane, bay.ridge_start, bay.ridge_end)
        return plane, axis_u

    def _resolve_member_width(self, family_name, type_name):
        member_width = self.get_type_width(family_name, type_name)
        if member_width is None or member_width <= 0.0:
            member_width = _member_width_from_text(type_name or "")
        if member_width is None or member_width <= 0.0:
            member_width = _member_width_from_text(family_name or "")
        if member_width is None or member_width <= 0.0:
            member_width = inches_to_feet(1.5)
        return member_width

    def _rafter_lines_with_stations(self, plan, field, bay):
        """Return ``[(ridge_station, eave_point, ridge_point), ...]`` for a field.

        The list is sorted by ridge_station so that both sides of a bay can be
        matched by position instead of by list index.  ``eave_point`` is the
        lower (support) end of each analytical rafter; ``ridge_point`` is the
        upper (ridge) end.
        """
        spacing = self.config.stud_spacing_ft
        if spacing <= 0.0 or field is None:
            return []

        sloped = _sloped_planes(getattr(plan.roof_info, "planes", []) or [])
        if field.plane_index < 0 or field.plane_index >= len(sloped):
            return []

        plane = sloped[field.plane_index]

        support_edges = list(
            getattr(field, "eave_edges", [])
            or getattr(field, "ledger_edges", [])
            or []
        )
        if not support_edges:
            return []

        axis_u, _ = _field_direction_axes(field, plane, bay.ridge_start, bay.ridge_end)

        result = []
        for ridge_point in _stations_along_segment(bay.ridge_start, bay.ridge_end, spacing):
            ridge_station = _ridge_station_on_segment(
                ridge_point,
                bay.ridge_start,
                bay.ridge_end,
            )
            candidate_edges = [
                edge for edge in support_edges
                if _segment_covers_ridge_station(
                    edge[0],
                    edge[1],
                    bay.ridge_start,
                    bay.ridge_end,
                    ridge_station,
                )
            ]
            if not candidate_edges:
                continue

            eave_point = None

            # Prefer a support intersection along the local rafter axis.
            if axis_u is not None:
                eave_point = _support_point_along_local_axis(
                    ridge_point,
                    candidate_edges,
                    plane,
                    axis_u,
                )

            # Fallback: closest valid support projection by ridge station.
            if eave_point is None:
                eave_point = _project_to_best_support(
                    ridge_point,
                    candidate_edges,
                    bay.ridge_start,
                    bay.ridge_end,
                )

            if eave_point is None:
                continue
            if _dist(eave_point, ridge_point) < MIN_MEMBER_LENGTH:
                continue

            result.append((ridge_station, eave_point, ridge_point))

        result.sort(key=lambda t: t[0])
        return result

    def _make_rafter_ties_from_plan(self, plan):
        """Create horizontal ties that intersect opposing rafters at one height.

        The tie endpoints are solved directly on the two analytical rafter
        segments for each station-matched pair, so each endpoint stays attached
        to its rafter centerline instead of being flattened after interpolation.
        """
        from Autodesk.Revit.DB import XYZ

        members = []
        seen = set()
        host_kind = getattr(plan.roof_info, "kind", "roof")
        host_id = getattr(plan.roof_info, "element_id", None)
        spacing = self.config.stud_spacing_ft
        if spacing <= 0.0:
            return members

        rafter_width = self._resolve_member_width(
            self.config.stud_family_name,
            self.config.stud_type_name,
        )
        # Beam centerlines should stop at opposing rafter faces, not centers.
        trim_each_end = rafter_width * 0.5
        station_tolerance = max(spacing * 0.001, 1e-4)
        max_station_drift = spacing * 0.25

        for bay in getattr(plan, "bays", []) or []:
            field_a = self._field_for_bay_side(plan, bay.index, bay.plane_a_index)
            field_b = self._field_for_bay_side(plan, bay.index, bay.plane_b_index)
            if field_a is None or field_b is None:
                continue

            # Skip hip-end bays.  A triangular hip field has no eave edge
            # (the eave collapses to a corner point), so there are no opposing
            # pairs to connect.
            if (not getattr(field_a, "eave_edges", [])
                    or not getattr(field_b, "eave_edges", [])):
                continue

            lines_a = self._rafter_lines_with_stations(plan, field_a, bay)
            lines_b = self._rafter_lines_with_stations(plan, field_b, bay)
            if not lines_a or not lines_b:
                continue

            for line_a, line_b in _match_lines_by_station(lines_a, lines_b, station_tolerance):
                station_a, eave_a, ridge_a = line_a
                station_b, eave_b, ridge_b = line_b

                z_low, z_high = _common_z_interval(eave_a, ridge_a, eave_b, ridge_b)
                if (z_high - z_low) < EDGE_TOL:
                    continue

                desired_a = eave_a.Z + ((ridge_a.Z - eave_a.Z) * RAFTER_TIE_FRACTION)
                desired_b = eave_b.Z + ((ridge_b.Z - eave_b.Z) * RAFTER_TIE_FRACTION)
                tie_z = (desired_a + desired_b) / 2.0
                if tie_z < z_low:
                    tie_z = z_low
                elif tie_z > z_high:
                    tie_z = z_high

                tie_a = _point_on_segment_at_z(eave_a, ridge_a, tie_z)
                tie_b = _point_on_segment_at_z(eave_b, ridge_b, tie_z)
                if tie_a is None or tie_b is None:
                    continue

                tie_station_a = _ridge_station_on_segment(tie_a, bay.ridge_start, bay.ridge_end)
                tie_station_b = _ridge_station_on_segment(tie_b, bay.ridge_start, bay.ridge_end)
                if abs(tie_station_a - tie_station_b) > max_station_drift:
                    continue

                tie_a, tie_b = _trim_segment_ends(tie_a, tie_b, trim_each_end)
                if _dist(tie_a, tie_b) < MIN_MEMBER_LENGTH:
                    continue

                if abs(tie_a.Z - tie_b.Z) > 1e-6:
                    tie_plane_z = (tie_a.Z + tie_b.Z) / 2.0
                    tie_a = XYZ(tie_a.X, tie_a.Y, tie_plane_z)
                    tie_b = XYZ(tie_b.X, tie_b.Y, tie_plane_z)

                key = (_pt_key(tie_a), _pt_key(tie_b))
                if key in seen or (_pt_key(tie_b), _pt_key(tie_a)) in seen:
                    continue
                seen.add(key)

                member = FramingMember(FramingMember.HEADER, tie_a, tie_b)
                member.member_type = "RAFTER_TIE"
                member.family_name = self.config.stud_family_name
                member.type_name = self.config.stud_type_name
                member.rotation = 0.0
                member.host_kind = host_kind
                member.host_id = host_id
                members.append(member)

        return members

    # ------------------------------------------------------------------
    # Dead-code methods removed:
    #   _make_collar_ties_from_plan  – collar ties are only correct when
    #     there is NO ridge board; V2 always assumes a ridge board.
    #   _make_ceiling_joists_from_plan / kickers – ceiling framing belongs
    #     to the dedicated ceiling framing tool and kickers require a
    #     ceiling-joist reference that the roof tool does not have.
    # ------------------------------------------------------------------

    def place_members(self, members, host_info):
        from System.Collections.Generic import List
        from Autodesk.Revit.DB import (
            Curve, Line, SketchPlane, Plane, BeamSystem,
            LayoutRuleFixedDistance, BeamSystemJustifyType,
        )
        from wf_families import activate_symbol
        try:
            from Autodesk.Revit.DB.Structure import StructuralFramingUtils
        except Exception:
            StructuralFramingUtils = None

        placed_elements = []

        # 1. Split members
        regular_members = [m for m in members if getattr(m, "member_type", "") != "BEAM_SYSTEM_V2"]
        bs_members = [m for m in members if getattr(m, "member_type", "") == "BEAM_SYSTEM_V2"]

        # 2. Place regular members (ridges, fascias)
        created_regular = BaseFramingEngine.place_members(self, regular_members, host_info)
        if created_regular:
            placed_elements.extend(created_regular)

        # 3. Create BeamSystems
        for member in bs_members:
            plane = member.plane
            if plane is None:
                continue

            # Convert loop
            world_pts = []
            for lx, ly in member.field_loop_local:
                # Cancel the host depth offset so the profile sits on plane.origin.
                world_pts.append(
                    plane.point_at(lx, ly, -plane.target_layer_depth)
                )

            # Filter consecutive duplicate points to maintain loop contiguity
            valid_pts = []
            for p in world_pts:
                if not valid_pts:
                    valid_pts.append(p)
                elif _dist(p, valid_pts[-1]) > 0.005:
                    valid_pts.append(p)

            # Ensure the start and end points aren't duplicated
            if len(valid_pts) > 1 and _dist(valid_pts[0], valid_pts[-1]) <= 0.005:
                valid_pts.pop()

            if len(valid_pts) < 3:
                continue

            profile = List[Curve]()
            invalid_profile = False
            for i in range(len(valid_pts)):
                p1 = valid_pts[i]
                p2 = valid_pts[(i + 1) % len(valid_pts)]
                if _dist(p1, p2) <= 0.005:
                    invalid_profile = True
                    break
                profile.Add(Line.CreateBound(p1, p2))

            if invalid_profile or profile.Count < 3:
                continue

            # create SketchPlane
            revit_plane = Plane.CreateByNormalAndOrigin(plane.normal, plane.origin)
            sketch_plane = SketchPlane.Create(self.doc, revit_plane)

            # direction
            dir_x, dir_y = member.direction_local
            pt0 = plane.point_at(0.0, 0.0, -plane.target_layer_depth)
            pt1 = plane.point_at(dir_x, dir_y, -plane.target_layer_depth)
            b_dir = _normalize(pt1 - pt0)
            if b_dir is None:
                continue

            bs = BeamSystem.Create(self.doc, profile, sketch_plane, b_dir, False)
            symbol = self._resolve_symbol(member)
            if symbol:
                activate_symbol(self.doc, symbol)
                bs.BeamType = symbol
            bs.LayoutRule = LayoutRuleFixedDistance(
                member.spacing,
                BeamSystemJustifyType.DirectionLine,
            )

            # Drop the system from the roof finish plane down to the top of the structural target layer.
            _set_beam_system_elevation(bs, -_resolve_roof_layer_top_depth(plane))
            try:
                tag_instance(bs, host_info, member)
            except Exception:
                pass

            placed_elements.append(bs)

            self.doc.Regenerate()
            rafter_ids = bs.GetBeamIds()
            rafters = [self.doc.GetElement(i) for i in rafter_ids]

            self._set_coping_distance_zero(created_regular)
            self._set_coping_distance_zero(rafters)

            # Apply coping
            for r in rafters:
                if StructuralFramingUtils:
                    try:
                        StructuralFramingUtils.DisallowJoinAtEnd(r, 0)
                        StructuralFramingUtils.DisallowJoinAtEnd(r, 1)
                    except Exception:
                        pass

                try:
                    tag_instance(r, host_info, member)
                except Exception:
                    pass

                try:
                    apply_bom_metadata(r, host_info, "RAFTER")
                except Exception:
                    pass

                add_coping = getattr(r, "AddCoping", None)
                if add_coping is None:
                    placed_elements.append(r)
                    continue

                for rb_inst in created_regular:
                    if not self._elements_are_near(r, rb_inst):
                        continue
                    try:
                        add_coping(rb_inst)
                    except Exception:
                        pass

                placed_elements.append(r)

        return placed_elements

    @staticmethod
    def _set_coping_distance_zero(instances):
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

    @staticmethod
    def _elements_are_near(first, second, tolerance=0.25):
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

    def _resolve_board_center_depth(self, plane, family_name, type_name):
        layer_top_depth = _resolve_roof_layer_top_depth(plane)
        member_depth = self.get_type_depth(family_name, type_name)
        if member_depth is None or member_depth <= 0.0:
            member_depth = _member_depth_from_text(type_name or "")
        if member_depth is None or member_depth <= 0.0:
            member_depth = _member_depth_from_text(family_name or "")
        if member_depth is None or member_depth <= 0.0:
            member_depth = inches_to_feet(1.5)
        return layer_top_depth + (member_depth / 2.0)

    def _resolve_member_center_depth(self, plane, family_name, type_name):
        layer_depth = float(getattr(plane, "target_layer_depth", 0.0) or 0.0)
        member_depth = self.get_type_depth(family_name, type_name)
        if member_depth is None or member_depth <= 0.0:
            member_depth = _member_depth_from_text(type_name or "")
        if member_depth is None or member_depth <= 0.0:
            member_depth = _member_depth_from_text(family_name or "")
        if member_depth is None or member_depth <= 0.0:
            member_depth = inches_to_feet(1.5)
        return layer_depth + (member_depth / 2.0)
