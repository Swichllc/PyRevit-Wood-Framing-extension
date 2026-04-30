# -*- coding: utf-8 -*-
"""Two-wall cleanup for corner and T wall framing assemblies."""

import math

from wf_wall_framing_v4 import (
    DEFAULT_LUMBER_DEPTH,
    MIN_MEMBER_LENGTH,
    PLATE_ROTATION,
    PLATE_THICKNESS,
    STUD_THICKNESS,
    WallCavityFramingV4Engine,
    _side_stud_position,
    _vertical_bounds,
)
from wf_geometry import inches_to_feet
from wf_tracking import get_tracking_data


JOIN_KIND_CORNER = "corner"
JOIN_KIND_T = "t_intersection"

STYLE_CORNER_INSULATED = "corner_insulated"
STYLE_CORNER_CAVITY = "corner_cavity"
STYLE_T_BLOCKING_NAILER = "t_blocking_nailer"
STYLE_T_ASSEMBLY = "t_assembly"

JOIN_TOL = inches_to_feet(4.0)
ANGLE_DOT_TOL = 0.25
DEFAULT_DELETE_RADIUS = inches_to_feet(24.0)

DELETE_MEMBER_ROLES = set([
    "CORNER_STUD",
    "CORNER_BACKING_STUD",
    "CORNER_RETURN_STUD",
    "CORNER_BLOCKING",
    "T_BRANCH_STUD",
    "T_BRANCH_BACKING_STUD",
    "T_BACKING_STUD",
    "T_NAILER",
    "T_BLOCKING",
])

DELETE_TRACKING_KINDS = set(["wall_join"])


class WallJoinCleanupError(Exception):
    pass


class WallJoinRelation(object):
    def __init__(self):
        self.kind = None
        self.point = None
        self.angle_degrees = 0.0
        self.hosts = []

        self.owner_host = None
        self.owner_end = None
        self.secondary_host = None
        self.secondary_end = None

        self.main_host = None
        self.main_distance = None
        self.branch_host = None
        self.branch_end = None


class WallJoinCleanupResult(object):
    def __init__(self):
        self.join_kind = None
        self.style_key = None
        self.angle_degrees = 0.0
        self.deleted_count = 0
        self.requested_count = 0
        self.placed_count = 0
        self.skipped_count = 0
        self.warnings = []


class WallJoinAssemblyPlan(object):
    def __init__(self):
        self.join_kind = None
        self.style_key = None
        self.angle_degrees = 0.0
        self.hosts = []
        self.members = []


def analyze_wall_join(doc, walls, config):
    """Analyze the two selected walls and return their join relationship."""
    if walls is None or len(walls) != 2:
        raise WallJoinCleanupError("Select exactly two walls.")

    engine = WallCavityFramingV4Engine(doc, config)
    hosts = _analyze_hosts(engine, walls, use_raw_face_domain=False)
    relation = _classify_hosts(hosts[0], hosts[1])
    relation.hosts = hosts
    return relation


def cleanup_selected_wall_join(doc, walls, config, style_key, delete_radius=None):
    """Replace previous join assembly members and place a chosen assembly."""
    if delete_radius is None:
        delete_radius = DEFAULT_DELETE_RADIUS

    engine = WallCavityFramingV4Engine(doc, config)
    detection_hosts = _analyze_hosts(engine, walls, use_raw_face_domain=False)
    placement_hosts = _analyze_hosts(engine, walls, use_raw_face_domain=True)
    detection_relation = _classify_hosts(detection_hosts[0], detection_hosts[1])
    relation = _rebind_relation_to_hosts(detection_relation, placement_hosts)

    result = WallJoinCleanupResult()
    result.join_kind = relation.kind
    result.style_key = style_key
    result.angle_degrees = relation.angle_degrees

    result.deleted_count = _delete_tracked_members_near_join(
        doc,
        walls,
        relation.point,
        delete_radius,
    )

    members = _build_join_members(engine, relation, style_key)
    members = _dedupe_members(members)
    result.requested_count = len(members)

    placed_count = 0
    for host in placement_hosts:
        host_id_text = _element_id_text(host.element_id)
        host_members = [
            member for member in members
            if _element_id_text(getattr(member, "host_id", None)) == host_id_text
        ]
        if not host_members:
            continue
        placed = engine.place_members(host_members, host)
        placed_count += len(placed)

    result.placed_count = placed_count
    result.skipped_count = max(0, result.requested_count - result.placed_count)
    if result.skipped_count:
        result.warnings.append(
            "{0} requested join member(s) were rejected by geometry validation or family placement.".format(
                result.skipped_count
            )
        )
    return result


def build_wall_join_assembly_plans(doc, walls, config, corner_style_key,
                                   t_style_key, detection_hosts=None):
    """Build join assemblies for selected walls without deleting wall studs.

    Selected walls are paired with other walls in the document so rerunning
    Wall Framing for one wall can recreate that wall's side of a prior join
    assembly. Members owned by unselected walls are not returned.
    """
    if walls is None or len(walls) < 1:
        return []

    target_ids = set()
    for wall in walls:
        wall_id = _element_id_text(getattr(wall, "Id", None))
        if wall_id:
            target_ids.add(wall_id)
    if not target_ids:
        return []

    engine = WallCavityFramingV4Engine(doc, config)
    context_walls = _join_context_walls(walls)
    if detection_hosts is None:
        detection_hosts = _analyze_host_map(engine, context_walls, False)
    placement_hosts = {}

    plans = []
    seen_pairs = set()
    seen_members = set()

    for wall in walls:
        wall_id = _element_id_text(getattr(wall, "Id", None))
        if wall_id is None or wall_id not in detection_hosts:
            continue

        host_a = detection_hosts.get(wall_id)
        for other in context_walls:
            other_id = _element_id_text(getattr(other, "Id", None))
            if other_id is None or other_id == wall_id:
                continue
            pair_key = tuple(sorted([wall_id, other_id]))
            if pair_key in seen_pairs:
                continue
            seen_pairs.add(pair_key)

            host_b = detection_hosts.get(other_id)
            if host_b is None:
                continue

            try:
                detection_relation = _classify_hosts(host_a, host_b)
            except WallJoinCleanupError:
                continue

            style_key = None
            if detection_relation.kind == JOIN_KIND_CORNER:
                style_key = corner_style_key
            elif detection_relation.kind == JOIN_KIND_T:
                style_key = t_style_key
            if not style_key:
                continue

            raw_a = _raw_host_for_wall(engine, wall, placement_hosts)
            raw_b = _raw_host_for_wall(engine, other, placement_hosts)
            if raw_a is None or raw_b is None:
                continue

            detection_relation.hosts = [host_a, host_b]
            try:
                relation = _rebind_relation_to_hosts(
                    detection_relation,
                    [raw_a, raw_b],
                )
                members = _build_join_members(engine, relation, style_key)
            except WallJoinCleanupError:
                continue

            selected_members = []
            for member in _dedupe_members(members):
                member_host_id = _element_id_text(getattr(member, "host_id", None))
                if member_host_id not in target_ids:
                    continue
                member_key = _member_key(member)
                if member_key in seen_members:
                    continue
                seen_members.add(member_key)
                selected_members.append(member)
            if not selected_members:
                continue

            plan = WallJoinAssemblyPlan()
            plan.join_kind = relation.kind
            plan.style_key = style_key
            plan.angle_degrees = relation.angle_degrees
            plan.hosts = [
                host for host in (raw_a, raw_b)
                if _element_id_text(getattr(host, "element_id", None)) in target_ids
            ]
            plan.members = selected_members
            plans.append(plan)

    return plans


def _join_context_walls(selected_walls):
    result = []
    seen = set()

    def add_wall(wall):
        wall_id = _element_id_text(getattr(wall, "Id", None))
        if wall is None or wall_id is None or wall_id in seen:
            return
        seen.add(wall_id)
        result.append(wall)

    for wall in selected_walls or []:
        add_wall(wall)
    return result


def _raw_host_for_wall(engine, wall, host_map):
    wall_id = _element_id_text(getattr(wall, "Id", None))
    if wall_id is None:
        return None
    if wall_id in host_map:
        return host_map.get(wall_id)
    try:
        host = engine._analyze_wall(wall, use_raw_face_domain=True)
    except Exception:
        host = None
    if host is not None:
        host.kind = "wall_join"
    host_map[wall_id] = host
    return host


def _analyze_host_map(engine, walls, use_raw_face_domain):
    result = {}
    for wall in walls or []:
        wall_id = _element_id_text(getattr(wall, "Id", None))
        if wall_id is None:
            continue
        try:
            host = engine._analyze_wall(
                wall,
                use_raw_face_domain=use_raw_face_domain,
            )
        except Exception:
            host = None
        if host is None:
            continue
        host.kind = "wall_join"
        result[wall_id] = host
    return result


def _analyze_hosts(engine, walls, use_raw_face_domain):
    hosts = []
    for wall in walls:
        host = engine._analyze_wall(wall, use_raw_face_domain=use_raw_face_domain)
        if host is None:
            raise WallJoinCleanupError(
                "Could not read wall side-face geometry for one selected wall."
            )
        host.kind = "wall_join"
        hosts.append(host)
    return hosts


def _rebind_relation_to_hosts(source, hosts):
    relation = WallJoinRelation()
    relation.kind = source.kind
    relation.point = source.point
    relation.angle_degrees = source.angle_degrees
    relation.hosts = hosts

    if source.kind == JOIN_KIND_CORNER:
        relation.owner_host = _matching_host(hosts, source.owner_host)
        relation.owner_end = source.owner_end
        relation.secondary_host = _matching_host(hosts, source.secondary_host)
        relation.secondary_end = source.secondary_end
        return relation

    if source.kind == JOIN_KIND_T:
        relation.main_host = _matching_host(hosts, source.main_host)
        relation.main_distance = _projection_distance(
            relation.main_host,
            source.point,
        )
        relation.branch_host = _matching_host(hosts, source.branch_host)
        relation.branch_end = source.branch_end
        return relation

    return relation


def _matching_host(hosts, source_host):
    source_id = _element_id_text(getattr(source_host, "element_id", None))
    for host in hosts:
        if _element_id_text(getattr(host, "element_id", None)) == source_id:
            return host
    raise WallJoinCleanupError("Could not match wall join analysis hosts.")


def _classify_hosts(host_a, host_b):
    relation = WallJoinRelation()
    relation.point = _line_intersection_xy(
        host_a.start_point,
        host_a.direction,
        host_b.start_point,
        host_b.direction,
    )
    if relation.point is None:
        raise WallJoinCleanupError("The selected walls are parallel or nearly parallel.")

    angle_dot = abs(host_a.direction.DotProduct(host_b.direction))
    if angle_dot > ANGLE_DOT_TOL:
        raise WallJoinCleanupError(
            "The selected walls are not close enough to a 90 degree join."
        )
    relation.angle_degrees = _angle_between_degrees(host_a.direction, host_b.direction)

    state_a, end_a, distance_a = _host_join_state(host_a, relation.point)
    state_b, end_b, distance_b = _host_join_state(host_b, relation.point)

    if state_a == "outside" or state_b == "outside":
        raise WallJoinCleanupError(
            "The selected walls do not meet within the current join tolerance."
        )

    if state_a == "end" and state_b == "end":
        relation.kind = JOIN_KIND_CORNER
        if _element_id_number(host_a.element_id) <= _element_id_number(host_b.element_id):
            relation.owner_host = host_a
            relation.owner_end = end_a
            relation.secondary_host = host_b
            relation.secondary_end = end_b
        else:
            relation.owner_host = host_b
            relation.owner_end = end_b
            relation.secondary_host = host_a
            relation.secondary_end = end_a
        return relation

    if state_a == "interior" and state_b == "end":
        relation.kind = JOIN_KIND_T
        relation.main_host = host_a
        relation.main_distance = distance_a
        relation.branch_host = host_b
        relation.branch_end = end_b
        return relation

    if state_b == "interior" and state_a == "end":
        relation.kind = JOIN_KIND_T
        relation.main_host = host_b
        relation.main_distance = distance_b
        relation.branch_host = host_a
        relation.branch_end = end_a
        return relation

    raise WallJoinCleanupError(
        "Only end-to-end corners and end-to-side T intersections are supported."
    )


def _build_join_members(engine, relation, style_key):
    if relation.kind == JOIN_KIND_CORNER:
        if style_key not in (STYLE_CORNER_INSULATED, STYLE_CORNER_CAVITY):
            raise WallJoinCleanupError("Choose a corner assembly for a corner join.")
        return _corner_members(engine, relation, style_key)

    if relation.kind == JOIN_KIND_T:
        if style_key not in (STYLE_T_BLOCKING_NAILER, STYLE_T_ASSEMBLY):
            raise WallJoinCleanupError("Choose a T assembly for a T intersection.")
        return _t_members(engine, relation, style_key)

    raise WallJoinCleanupError("Unsupported wall join type.")


def _corner_members(engine, relation, style_key):
    members = []
    owner = relation.owner_host
    secondary = relation.secondary_host

    owner_d = _end_stud_d(owner, relation.owner_end, secondary, relation.point)
    secondary_d = _end_stud_d(
        secondary,
        relation.secondary_end,
        owner,
        relation.point,
    )
    _add_vertical_member(engine, members, owner, "CORNER_STUD", owner_d, True)
    _add_vertical_member(engine, members, secondary, "CORNER_STUD", secondary_d, True)

    owner_backing_d = _inboard_stud_d(owner, relation.owner_end, owner_d, 1)
    owner_backing_placed = _add_vertical_member(
        engine,
        members,
        owner,
        "CORNER_BACKING_STUD",
        owner_backing_d,
        False,
    )

    if style_key == STYLE_CORNER_INSULATED:
        return members

    if style_key == STYLE_CORNER_CAVITY:
        secondary_backing_d = _inboard_stud_d(
            secondary,
            relation.secondary_end,
            secondary_d,
            1,
        )
        secondary_backing_placed = _add_vertical_member(
            engine,
            members,
            secondary,
            "CORNER_RETURN_STUD",
            secondary_backing_d,
            False,
        )
        if owner_backing_placed:
            _add_blocking_between(
                engine,
                members,
                owner,
                owner_d,
                owner_backing_d,
                "CORNER_BLOCKING",
            )
        if secondary_backing_placed:
            _add_blocking_between(
                engine,
                members,
                secondary,
                secondary_d,
                secondary_backing_d,
                "CORNER_BLOCKING",
            )
    return members


def _t_members(engine, relation, style_key):
    members = []
    main = relation.main_host
    branch = relation.branch_host

    main_d = _clamp_d(main, relation.main_distance)
    branch_d = _end_stud_d(branch, relation.branch_end, main, relation.point)
    _add_required_edge_stud(engine, members, branch, "T_BRANCH_STUD", branch_d)

    branch_backing_d = _inboard_stud_d(
        branch,
        relation.branch_end,
        branch_d,
        1,
    )

    if style_key == STYLE_T_ASSEMBLY:
        offset = _t_backing_offset(branch)
        _add_vertical_member(
            engine,
            members,
            branch,
            "T_BRANCH_BACKING_STUD",
            branch_backing_d,
            False,
        )
        _add_offset_vertical_member(
            engine,
            members,
            main,
            "T_BACKING_STUD",
            main_d,
            -1.0,
            offset,
        )
        _add_offset_vertical_member(
            engine,
            members,
            main,
            "T_BACKING_STUD",
            main_d,
            1.0,
            offset,
        )
    else:
        _add_vertical_member(
            engine,
            members,
            branch,
            "T_BRANCH_BACKING_STUD",
            branch_backing_d,
            False,
        )
        if _add_vertical_member(engine, members, main, "T_NAILER", main_d, False):
            half_span = max(
                inches_to_feet(8.0),
                min(inches_to_feet(12.0), engine.config.stud_spacing_ft * 0.5),
            )
            _add_blocking_between(
                engine,
                members,
                main,
                main_d - half_span,
                main_d + half_span,
                "T_BLOCKING",
            )
    return members


def _add_offset_vertical_member(engine, members, host, role, center_d, side_sign,
                                offset):
    side_sign = -1.0 if side_sign < 0.0 else 1.0
    offset = max(STUD_THICKNESS, float(offset or STUD_THICKNESS))
    distances = [
        offset,
        offset + STUD_THICKNESS,
        offset + STUD_THICKNESS * 2.0,
        max(STUD_THICKNESS * 0.5, offset - STUD_THICKNESS * 0.5),
    ]
    for distance in distances:
        d = center_d + side_sign * distance
        if not _valid_stud_center_d(host, d):
            continue
        if _has_vertical_member_near(members, host, d, STUD_THICKNESS * 0.5):
            continue
        if _add_vertical_member(engine, members, host, role, d, False):
            return True
    return False


def _add_vertical_member(engine, members, host, role, d, is_side):
    d = _clamp_d(host, d)
    validation_role = "SIDE_STUD" if is_side else "STUD"
    member = engine._vertical_member_at_d(host, validation_role, d, None, None, False)
    if member is None:
        return False
    member.member_type = role
    member.host_kind = "wall_join"
    member.host_id = host.element_id
    members.append(member)
    return True


def _add_required_edge_stud(engine, members, host, role, d):
    if _add_vertical_member(engine, members, host, role, d, True):
        return True
    return _add_vertical_member(engine, members, host, role, d, False)


def _valid_stud_center_d(host, d):
    if d is None:
        return False
    try:
        d = float(d)
    except Exception:
        return False
    return STUD_THICKNESS * 0.5 <= d <= host.length - STUD_THICKNESS * 0.5


def _has_vertical_member_near(members, host, d, tolerance):
    host_id = _element_id_text(getattr(host, "element_id", None))
    for member in members:
        if not getattr(member, "is_column", False):
            continue
        if _element_id_text(getattr(member, "host_id", None)) != host_id:
            continue
        current_d = _member_distance_on_host(host, member)
        if current_d is None:
            continue
        if abs(current_d - d) <= tolerance:
            return True
    return False


def _member_distance_on_host(host, member):
    point = getattr(member, "start_point", None)
    if point is None:
        return None
    try:
        return _projection_distance(host, point)
    except Exception:
        return None


def _add_blocking_between(engine, members, host, d0, d1, role):
    start_d = _clamp_d(host, min(d0, d1))
    end_d = _clamp_d(host, max(d0, d1))
    if end_d - start_d < MIN_MEMBER_LENGTH:
        return

    mid_d = (start_d + end_d) * 0.5
    for z_abs in _blocking_heights(engine, host, mid_d):
        member = _horizontal_member(engine, host, role, start_d, end_d, z_abs)
        if member is not None:
            members.append(member)


def _horizontal_member(engine, host, role, start_d, end_d, z_abs):
    family = engine.config.bottom_plate_family_name or engine.config.stud_family_name
    type_name = engine.config.bottom_plate_type_name or engine.config.stud_type_name
    depth = engine._wall_member_depth(host, family, type_name, False)
    start = host.point_at_abs(start_d, z_abs)
    end = host.point_at_abs(end_d, z_abs)
    member = engine._member_from_points(
        host,
        role,
        start,
        end,
        family,
        type_name,
        False,
        PLATE_ROTATION,
        PLATE_THICKNESS,
        depth,
    )
    if member is None:
        return None
    member.member_type = role
    member.host_kind = "wall_join"
    member.host_id = host.element_id
    return member


def _blocking_heights(engine, host, d):
    bounds = _vertical_bounds(host.outer_loop, _clamp_d(host, d))
    if bounds is None:
        return []
    bottom_z = bounds[0] + engine._stud_bottom()
    top_z = bounds[1] - engine._top_plate_stack()
    clear = top_z - bottom_z
    if clear < MIN_MEMBER_LENGTH * 2.0:
        return []
    if clear >= 7.0:
        factors = (0.25, 0.5, 0.75)
    elif clear >= 4.0:
        factors = (0.33, 0.67)
    else:
        factors = (0.5,)
    return [bottom_z + clear * factor for factor in factors]


def _delete_tracked_members_near_join(doc, walls, join_point, radius):
    from Autodesk.Revit.DB import BuiltInCategory, FilteredElementCollector

    wall_ids = set([_element_id_text(wall.Id) for wall in walls])
    delete_ids = []
    seen = set()
    for category in (
            BuiltInCategory.OST_StructuralFraming,
            BuiltInCategory.OST_StructuralColumns):
        try:
            collector = (
                FilteredElementCollector(doc)
                .OfCategory(category)
                .WhereElementIsNotElementType()
            )
        except Exception:
            continue
        for element in collector:
            tracking = get_tracking_data(element)
            if not tracking:
                continue
            if tracking.get("kind") not in DELETE_TRACKING_KINDS:
                continue
            if tracking.get("host") not in wall_ids:
                continue
            role = (tracking.get("member") or "").upper()
            if role not in DELETE_MEMBER_ROLES:
                continue
            if not _element_near_point_xy(element, join_point, radius):
                continue
            element_id = getattr(element, "Id", None)
            key = _element_id_text(element_id)
            if key in seen:
                continue
            seen.add(key)
            delete_ids.append(element_id)

    deleted = 0
    for element_id in delete_ids:
        try:
            doc.Delete(element_id)
            deleted += 1
        except Exception:
            pass
    return deleted


def _element_near_point_xy(element, point, radius):
    try:
        bbox = element.get_BoundingBox(None)
    except Exception:
        bbox = None
    if bbox is None or point is None:
        return False

    dx = _outside_interval_distance(point.X, bbox.Min.X, bbox.Max.X)
    dy = _outside_interval_distance(point.Y, bbox.Min.Y, bbox.Max.Y)
    return math.sqrt(dx * dx + dy * dy) <= radius


def _outside_interval_distance(value, low, high):
    if value < low:
        return low - value
    if value > high:
        return value - high
    return 0.0


def _line_intersection_xy(point_a, dir_a, point_b, dir_b):
    denom = dir_a.X * dir_b.Y - dir_a.Y * dir_b.X
    if abs(denom) < 1e-9:
        return None

    dx = point_b.X - point_a.X
    dy = point_b.Y - point_a.Y
    t = (dx * dir_b.Y - dy * dir_b.X) / denom

    from Autodesk.Revit.DB import XYZ
    return XYZ(
        point_a.X + dir_a.X * t,
        point_a.Y + dir_a.Y * t,
        point_a.Z,
    )


def _host_join_state(host, point):
    distance = _projection_distance(host, point)
    perp = _perpendicular_distance(host, point, distance)
    if perp > JOIN_TOL:
        return "outside", None, distance
    if distance < -JOIN_TOL or distance > host.length + JOIN_TOL:
        return "outside", None, distance
    if distance <= JOIN_TOL:
        return "end", 0, 0.0
    if distance >= host.length - JOIN_TOL:
        return "end", 1, host.length
    return "interior", None, distance


def _projection_distance(host, point):
    vec = point - host.start_point
    return vec.DotProduct(host.direction)


def _perpendicular_distance(host, point, distance):
    projected = host.start_point + host.direction.Multiply(distance)
    dx = point.X - projected.X
    dy = point.Y - projected.Y
    return math.sqrt(dx * dx + dy * dy)


def _angle_between_degrees(dir_a, dir_b):
    try:
        dot = max(-1.0, min(1.0, abs(dir_a.DotProduct(dir_b))))
        return math.degrees(math.acos(dot))
    except Exception:
        return 0.0


def _end_stud_d(host, end_index, other_host=None, join_point=None):
    if other_host is not None and join_point is not None:
        sign = _end_sign(end_index)
        core_half = _host_core_width(other_host) * 0.5
        stud_half = STUD_THICKNESS * 0.5
        center = _projection_distance(host, join_point) + sign * (core_half - stud_half)
        return _clamp_d(host, center)

    raw = 0.0 if end_index == 0 else host.length
    return _side_stud_position(host, raw)


def _inboard_stud_d(host, end_index, end_stud_d, stud_steps):
    offset = STUD_THICKNESS * max(1, int(stud_steps))
    if end_index == 0:
        return _clamp_d(host, end_stud_d + offset)
    return _clamp_d(host, end_stud_d - offset)


def _end_sign(end_index):
    return -1.0 if end_index == 0 else 1.0


def _host_core_width(host):
    target_layer = getattr(host, "target_layer", None)
    if target_layer is not None:
        width = float(getattr(target_layer, "width", 0.0) or 0.0)
        if width > STUD_THICKNESS:
            return width
    return DEFAULT_LUMBER_DEPTH


def _t_backing_offset(branch_host):
    width = 0.0
    target_layer = getattr(branch_host, "target_layer", None)
    if target_layer is not None:
        width = float(getattr(target_layer, "width", 0.0) or 0.0)
    return max(STUD_THICKNESS, width * 0.5 + STUD_THICKNESS * 0.5)


def _clamp_d(host, d):
    if d is None:
        return 0.0
    return max(0.0, min(host.length, float(d)))


def _dedupe_members(members):
    result = []
    seen = set()
    for member in members:
        key = _member_key(member)
        if key in seen:
            continue
        seen.add(key)
        result.append(member)
    return result


def _member_key(member):
    return (
        _element_id_text(getattr(member, "host_id", None)),
        getattr(member, "member_type", ""),
        _point_key(getattr(member, "start_point", None)),
        _point_key(getattr(member, "end_point", None)),
    )


def _point_key(point):
    if point is None:
        return None
    return (
        round(point.X, 5),
        round(point.Y, 5),
        round(point.Z, 5),
    )


def _element_id_number(element_id):
    value = _element_id_value(element_id)
    if value is None:
        return 0
    try:
        return int(value)
    except Exception:
        return 0


def _element_id_text(element_id):
    value = _element_id_value(element_id)
    if value is None:
        return None
    return str(value)


def _element_id_value(element_id):
    if element_id is None:
        return None
    try:
        integer_types = (int, long)
    except NameError:
        integer_types = (int,)
    if isinstance(element_id, integer_types):
        return element_id
    return getattr(element_id, "IntegerValue", getattr(element_id, "Value", None))
