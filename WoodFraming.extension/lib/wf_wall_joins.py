# -*- coding: utf-8 -*-
"""Wall-join helpers for corner and partition-backing stud placement."""

from wf_geometry import inches_to_feet
from wf_host import analyze_wall_host


END_JOIN_TOL = inches_to_feet(2.5)
INTERSECTION_TOL = inches_to_feet(1.0)
END_CLEARANCE = inches_to_feet(4.5)
PARALLEL_DOT_TOL = 0.98


class EndJoinPlan(object):
    """Stud placement plan for one wall end."""

    def __init__(self):
        self.has_join = False
        self.join_kind = None
        self.is_owner = True
        self.other_wall_id = None
        self.other_layer_width = 0.0
        self.inset = 0.0
        self.positions = []


class IntersectionJoinPlan(object):
    """Stud placement plan for a wall intersection along the wall run."""

    def __init__(self):
        self.distance = 0.0
        self.other_wall_id = None
        self.other_layer_width = 0.0
        self.positions = []


class WallJoinPlan(object):
    """All join-driven stud positions for a wall."""

    def __init__(self):
        self.ends = {0: EndJoinPlan(), 1: EndJoinPlan()}
        self.intersections = []


def build_wall_join_plan(doc, wall_info, config, stud_thickness):
    """Return corner and T-junction stud positions for a wall.

    Uses Revit's LocationCurve.ElementsAtJoin when available for end joins,
    then falls back to geometric matching for the rest.
    """
    plan = WallJoinPlan()
    if wall_info is None or wall_info.element is None:
        return plan

    this_id = _element_id_value(wall_info.element_id)
    joined_ids = _joined_wall_ids(wall_info.element)
    wall_start, wall_end = _framing_line_endpoints(wall_info)
    wall_ends = (wall_start, wall_end)

    try:
        from Autodesk.Revit.DB import FilteredElementCollector, Line, Wall

        walls = (
            FilteredElementCollector(doc)
            .OfClass(Wall)
            .WhereElementIsNotElementType()
        )
    except Exception:
        return plan

    for other in walls:
        other_id = _element_id_value(other.Id)
        if other_id == this_id:
            continue

        other_loc = other.Location
        if other_loc is None:
            continue
        other_curve = other_loc.Curve
        if not isinstance(other_curve, Line):
            continue

        other_start = other_curve.GetEndPoint(0)
        other_end = other_curve.GetEndPoint(1)
        other_dir = (other_end - other_start).Normalize()
        if _is_parallel(wall_info.direction, other_dir):
            continue

        try:
            other_info = analyze_wall_host(doc, other, config)
        except Exception:
            other_info = None
        if other_info is None:
            continue

        other_start, other_end = _framing_line_endpoints(other_info)

        for end_index in (0, 1):
            is_official_join = other_id in joined_ids[end_index]
            if joined_ids[end_index] and not is_official_join:
                continue
            other_end_index = _matching_end_index(
                wall_ends[end_index],
                other_start,
                other_end,
            )
            if other_end_index is None and not is_official_join:
                continue

            _merge_end_plan(
                plan.ends[end_index],
                this_id,
                other_id,
                other_info,
                stud_thickness,
                "corner" if other_end_index is not None else "termination",
            )

        for other_pt in (other_start, other_end):
            distance = _intersection_distance(wall_start, wall_info, other_pt)
            if distance is None:
                continue

            _merge_intersection_plan(
                plan.intersections,
                distance,
                other_id,
                other_info,
                stud_thickness,
            )

    for end_index in (0, 1):
        end_plan = plan.ends[end_index]
        if not end_plan.has_join:
            continue

        if end_plan.join_kind == "termination":
            end_plan.positions = _end_positions(
                wall_info.length,
                end_index,
                (0.0,),
            )
        elif end_plan.is_owner:
            end_plan.positions = _end_positions(
                wall_info.length,
                end_index,
                (0.0, stud_thickness),
            )
        else:
            inset = max(stud_thickness, end_plan.inset)
            end_plan.positions = _end_positions(
                wall_info.length,
                end_index,
                (inset,),
            )

    return plan


def _merge_end_plan(end_plan, this_id, other_id, other_info, stud_thickness, join_kind):
    width = _layer_width(other_info, stud_thickness)
    inset = max(stud_thickness, (width / 2.0) + (stud_thickness / 2.0))

    if not end_plan.has_join:
        end_plan.is_owner = True

    end_plan.has_join = True
    if _join_kind_rank(join_kind) > _join_kind_rank(end_plan.join_kind):
        end_plan.join_kind = join_kind

    if end_plan.join_kind == "corner":
        end_plan.is_owner = end_plan.is_owner and (this_id <= other_id)
    else:
        end_plan.is_owner = False

    if end_plan.other_wall_id is None or other_id < end_plan.other_wall_id:
        end_plan.other_wall_id = other_id
    if width > end_plan.other_layer_width:
        end_plan.other_layer_width = width
    if inset > end_plan.inset:
        end_plan.inset = inset


def _join_kind_rank(join_kind):
    if join_kind == "corner":
        return 2
    if join_kind == "termination":
        return 1
    return 0


def _merge_intersection_plan(items, distance, other_id, other_info, stud_thickness):
    width = _layer_width(other_info, stud_thickness)
    offset = max(stud_thickness, (width / 2.0) + (stud_thickness / 2.0))

    target = None
    for item in items:
        if abs(item.distance - distance) < stud_thickness:
            target = item
            break

    if target is None:
        target = IntersectionJoinPlan()
        target.distance = distance
        items.append(target)

    if target.other_wall_id is None or other_id < target.other_wall_id:
        target.other_wall_id = other_id
    if width > target.other_layer_width:
        target.other_layer_width = width

    offset = max(
        stud_thickness,
        (target.other_layer_width / 2.0) + (stud_thickness / 2.0),
    )
    target.positions = [distance - offset, distance + offset]


def _intersection_distance(start_pt, wall_info, point):
    distance = _projection_distance(start_pt, wall_info.direction, point)
    if distance <= END_CLEARANCE or distance >= wall_info.length - END_CLEARANCE:
        return None

    perp = _perpendicular_distance(start_pt, wall_info.direction, point)
    if perp > INTERSECTION_TOL:
        return None

    return distance


def _matching_end_index(target, start_pt, end_pt):
    if _xy_distance(target, start_pt) <= END_JOIN_TOL:
        return 0
    if _xy_distance(target, end_pt) <= END_JOIN_TOL:
        return 1
    return None


def _end_positions(length, end_index, offsets):
    positions = []
    for offset in offsets:
        pos = offset if end_index == 0 else length - offset
        if pos < -1e-9 or pos > length + 1e-9:
            continue
        if not _contains_close(positions, pos):
            positions.append(pos)
    return positions


def _contains_close(values, target):
    for value in values:
        if abs(value - target) < 1e-6:
            return True
    return False


def _joined_wall_ids(wall):
    result = {0: set(), 1: set()}
    loc = getattr(wall, "Location", None)
    if loc is None:
        return result

    for end_index in (0, 1):
        joined = None
        try:
            joined = loc.get_ElementsAtJoin(end_index)
        except Exception:
            try:
                joined = loc.ElementsAtJoin[end_index]
            except Exception:
                joined = None
        if joined is None:
            continue

        for elem in _iter_joined_elements(joined):
            elem_id = getattr(elem, "Id", None)
            if elem_id is None:
                continue
            result[end_index].add(_element_id_value(elem_id))

    return result


def _iter_joined_elements(joined):
    try:
        for elem in joined:
            yield elem
        return
    except Exception:
        pass

    try:
        count = joined.Size
    except Exception:
        count = 0
    for index in range(count):
        try:
            yield joined[index]
        except Exception:
            continue


def _layer_width(wall_info, stud_thickness):
    target_layer = getattr(wall_info, "target_layer", None)
    if target_layer is not None and target_layer.width > 1e-9:
        return target_layer.width

    base_info = getattr(wall_info, "wall_info", None)
    width = getattr(base_info, "width", 0.0)
    if width > 1e-9:
        return width

    return stud_thickness


def _framing_line_endpoints(wall_info):
    offset = getattr(wall_info, "target_layer_offset", 0.0)
    start = _shift_point(wall_info.start_point, wall_info.normal, offset)
    end = _shift_point(wall_info.end_point, wall_info.normal, offset)
    return start, end


def _element_id_value(element_id):
    return getattr(element_id, "Value", getattr(element_id, "IntegerValue", element_id))


def _shift_point(point, normal, offset):
    if point is None or normal is None or abs(offset) < 1e-9:
        return point
    return point + normal.Multiply(offset)


def _projection_distance(start_pt, direction, point):
    vec = point - start_pt
    return vec.DotProduct(direction)


def _perpendicular_distance(start_pt, direction, point):
    distance = _projection_distance(start_pt, direction, point)
    projected = start_pt + direction.Multiply(distance)
    return _xy_distance(projected, point)


def _is_parallel(dir_a, dir_b):
    try:
        return abs(dir_a.DotProduct(dir_b)) >= PARALLEL_DOT_TOL
    except Exception:
        return False


def _xy_distance(a, b):
    dx = a.X - b.X
    dy = a.Y - b.Y
    return (dx * dx + dy * dy) ** 0.5