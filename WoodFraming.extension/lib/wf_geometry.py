# -*- coding: utf-8 -*-
"""Geometry utilities for wall/floor/roof analysis.

All distance values are in FEET (Revit internal units) unless noted.
"""

import math


def inches_to_feet(inches):
    """Convert inches to feet."""
    return inches / 12.0


def feet_to_inches(feet):
    """Convert feet to inches."""
    return feet * 12.0


class WallInfo(object):
    """Extracted geometry and properties of a Revit wall."""

    def __init__(self):
        self.wall = None            # Revit Wall element
        self.wall_id = None
        self.start_point = None     # XYZ
        self.end_point = None       # XYZ
        self.direction = None       # XYZ unit vector along wall
        self.normal = None          # XYZ wall face normal
        self.length = 0.0           # feet
        self.height = 0.0           # feet (nominal / unconnected)
        self.width = 0.0            # feet
        self.base_offset = 0.0      # feet from level
        self.level_id = None
        self.level_elevation = 0.0  # feet
        self.location_line = None   # Revit WallLocationLine enum value
        # Lateral shift from the wall's current location line to core center.
        self.core_centerline_offset = 0.0
        self.is_straight = True
        self.angle = 0.0            # radians – wall direction in XY
        self.start_height = 0.0     # actual height at wall start
        self.end_height = 0.0       # actual height at wall end
        self.is_sloped_top = False  # True when start_height != end_height


class OpeningInfo(object):
    """Information about an opening (door/window) in a wall."""

    def __init__(self):
        self.element = None         # Revit FamilyInstance
        self.element_id = None
        self.is_door = False
        self.is_window = False
        self.center_point = None    # XYZ insertion point
        self.width = 0.0            # feet
        self.height = 0.0           # feet
        self.sill_height = 0.0     # feet from wall base (0 for doors)
        self.head_height = 0.0     # feet from wall base
        # Position along wall (distance from wall start to opening center)
        self.distance_along_wall = 0.0
        # Edges along wall (distance from wall start)
        self.left_edge = 0.0
        self.right_edge = 0.0


class FramingMember(object):
    """Describes a single framing member to be placed."""

    # Member types
    STUD = "stud"
    KING_STUD = "king_stud"
    JACK_STUD = "jack_stud"
    CRIPPLE_STUD = "cripple_stud"
    BOTTOM_PLATE = "bottom_plate"
    TOP_PLATE = "top_plate"
    HEADER = "header"
    SILL_PLATE = "sill_plate"

    def __init__(self, member_type, start_pt, end_pt):
        self.member_type = member_type
        self.start_point = start_pt   # XYZ
        self.end_point = end_pt       # XYZ
        self.family_name = None
        self.type_name = None
        self.rotation = 0.0           # radians, cross-section rotation
        self.is_column = False        # True → StructuralType.Column
        self.host_kind = None
        self.host_id = None
        self.layer_index = None


def analyze_wall(doc, wall):
    """Extract geometry info from a Revit Wall element.

    Args:
        doc: Revit Document
        wall: Autodesk.Revit.DB.Wall

    Returns:
        WallInfo or None if wall is not a simple straight wall.
    """
    from Autodesk.Revit.DB import Line as RvtLine

    loc = wall.Location
    if loc is None:
        return None

    curve = loc.Curve
    if not isinstance(curve, RvtLine):
        # Curved walls not supported in v1
        return None

    info = WallInfo()
    info.wall = wall
    info.wall_id = wall.Id
    info.start_point = curve.GetEndPoint(0)
    info.end_point = curve.GetEndPoint(1)
    info.length = curve.Length

    # Direction vector along wall
    dx = info.end_point.X - info.start_point.X
    dy = info.end_point.Y - info.start_point.Y
    dz = info.end_point.Z - info.start_point.Z
    length = math.sqrt(dx * dx + dy * dy + dz * dz)
    if length < 1e-9:
        return None
    info.direction = _make_xyz(dx / length, dy / length, dz / length)
    info.angle = math.atan2(info.direction.Y, info.direction.X)

    # Wall normal (perpendicular, horizontal)
    info.normal = safe_wall_normal(wall, info.direction)
    if info.normal is None:
        return None

    # Wall dimensions
    from Autodesk.Revit.DB import BuiltInParameter
    height_param = wall.get_Parameter(BuiltInParameter.WALL_USER_HEIGHT_PARAM)
    if height_param:
        info.height = height_param.AsDouble()
    else:
        info.height = 8.0  # fallback 8 ft

    info.width = wall.Width  # total wall width in feet

    _set_wall_location_data(wall, info)

    base_offset_param = wall.get_Parameter(BuiltInParameter.WALL_BASE_OFFSET)
    if base_offset_param:
        info.base_offset = base_offset_param.AsDouble()

    info.level_id = wall.LevelId
    level = doc.GetElement(wall.LevelId)
    if level:
        info.level_elevation = level.Elevation

    # Default: flat top
    info.start_height = info.height
    info.end_height = info.height

    # Detect sloped / vaulted wall top from actual geometry
    _detect_wall_slope(wall, info)

    return info


def find_openings(doc, wall, wall_info):
    """Find all door/window openings hosted in a wall.

    Args:
        doc: Revit Document
        wall: Autodesk.Revit.DB.Wall
        wall_info: WallInfo with wall geometry

    Returns:
        List of OpeningInfo sorted by distance along wall.
    """
    from Autodesk.Revit.DB import (
        BuiltInCategory,
        BuiltInParameter,
        FamilyInstance,
        Line as RvtLine,
        Opening as RvtOpening,
        Wall as RvtWall,
        WallKind,
    )

    openings = []
    host_base = wall_info.level_elevation + wall_info.base_offset

    try:
        insert_ids = wall.FindInserts(True, False, True, False)
    except Exception:
        insert_ids = []

    for eid in insert_ids:
        elem = doc.GetElement(eid)
        if elem is None:
            continue

        if isinstance(elem, FamilyInstance):
            cat = elem.Category
            if cat is None:
                continue

            host = getattr(elem, "Host", None)
            if host is None or host.Id != wall.Id:
                continue

            cat_id = getattr(cat.Id, "IntegerValue", getattr(cat.Id, "Value", None))
            is_door = (cat_id == int(BuiltInCategory.OST_Doors))
            is_window = (cat_id == int(BuiltInCategory.OST_Windows))
            if not (is_door or is_window):
                continue

            loc = elem.Location
            center_point = getattr(loc, "Point", None)
            if center_point is None:
                continue

            width = _get_opening_width(elem)
            height = _get_opening_height(elem)
            sill_height = _get_sill_height(elem) if is_window else 0.0
            head_height = sill_height + height
            center_dist = _project_point_on_wall(center_point, wall_info)

            _append_opening_info(
                openings,
                wall_info,
                elem,
                center_dist - (width / 2.0),
                center_dist + (width / 2.0),
                sill_height,
                head_height,
                is_window,
                center_point,
            )
            continue

        if isinstance(elem, RvtOpening):
            boundary = elem.BoundaryRect
            if not boundary or len(boundary) < 2:
                continue

            min_pt = boundary[0]
            max_pt = boundary[1]
            left_edge = min(
                _project_point_on_wall(min_pt, wall_info),
                _project_point_on_wall(max_pt, wall_info),
            )
            right_edge = max(
                _project_point_on_wall(min_pt, wall_info),
                _project_point_on_wall(max_pt, wall_info),
            )
            sill_height = min(min_pt.Z, max_pt.Z) - host_base
            head_height = max(min_pt.Z, max_pt.Z) - host_base

            _append_opening_info(
                openings,
                wall_info,
                elem,
                left_edge,
                right_edge,
                sill_height,
                head_height,
                sill_height > inches_to_feet(1.5),
            )
            continue

        if isinstance(elem, RvtWall):
            try:
                if elem.WallType.Kind != WallKind.Curtain:
                    continue
            except Exception:
                continue

            loc = elem.Location
            curve = getattr(loc, "Curve", None)
            if not isinstance(curve, RvtLine):
                continue

            left_edge = min(
                _project_point_on_wall(curve.GetEndPoint(0), wall_info),
                _project_point_on_wall(curve.GetEndPoint(1), wall_info),
            )
            right_edge = max(
                _project_point_on_wall(curve.GetEndPoint(0), wall_info),
                _project_point_on_wall(curve.GetEndPoint(1), wall_info),
            )
            if right_edge - left_edge < 1e-6:
                continue

            curtain_base = _wall_base_elevation(doc, elem)
            height_param = elem.get_Parameter(BuiltInParameter.WALL_USER_HEIGHT_PARAM)
            curtain_height = (
                height_param.AsDouble()
                if height_param is not None and height_param.HasValue
                else wall_info.height
            )
            sill_height = curtain_base - host_base
            head_height = sill_height + curtain_height

            _append_opening_info(
                openings,
                wall_info,
                elem,
                left_edge,
                right_edge,
                sill_height,
                head_height,
                sill_height > inches_to_feet(1.5),
            )

    # Sort by position along wall
    openings.sort(key=lambda o: o.distance_along_wall)
    return openings


def _project_point_on_wall(point, wall_info):
    """Project a 3D point onto the wall line. Returns distance from start."""
    dx = point.X - wall_info.start_point.X
    dy = point.Y - wall_info.start_point.Y
    # Dot product with wall direction (ignore Z)
    dist = dx * wall_info.direction.X + dy * wall_info.direction.Y
    return max(0.0, min(dist, wall_info.length))


def _append_opening_info(openings, wall_info, element, left_edge, right_edge,
                         sill_height, head_height, is_window,
                         center_point=None):
    """Add an opening if it does not duplicate an existing span."""
    left_edge = max(0.0, min(left_edge, wall_info.length))
    right_edge = max(0.0, min(right_edge, wall_info.length))
    if right_edge - left_edge < 1e-6:
        return
    if _has_duplicate_opening(openings, left_edge, right_edge):
        return

    oi = OpeningInfo()
    oi.element = element
    oi.element_id = getattr(element, "Id", None)
    oi.is_window = is_window
    oi.is_door = not is_window
    oi.left_edge = left_edge
    oi.right_edge = right_edge
    oi.distance_along_wall = (left_edge + right_edge) / 2.0
    oi.width = right_edge - left_edge
    oi.sill_height = max(0.0, sill_height)
    oi.head_height = max(oi.sill_height, head_height)
    oi.height = oi.head_height - oi.sill_height
    if center_point is None:
        oi.center_point = point_on_wall(
            wall_info,
            oi.distance_along_wall,
            oi.sill_height,
        )
    else:
        oi.center_point = center_point
    openings.append(oi)


def _has_duplicate_opening(openings, left_edge, right_edge, tol=None):
    """Return True when an opening span already exists in the list."""
    if tol is None:
        tol = inches_to_feet(1.5)
    for opening in openings:
        if (abs(opening.left_edge - left_edge) < tol and
                abs(opening.right_edge - right_edge) < tol):
            return True
    return False


def _wall_base_elevation(doc, wall):
    """Return a wall base elevation in model coordinates."""
    from Autodesk.Revit.DB import BuiltInParameter

    base_elevation = 0.0
    level = doc.GetElement(wall.LevelId)
    if level is not None:
        base_elevation = level.Elevation

    offset_param = wall.get_Parameter(BuiltInParameter.WALL_BASE_OFFSET)
    if offset_param is not None and offset_param.HasValue:
        base_elevation += offset_param.AsDouble()
    return base_elevation


def _get_opening_width(family_instance):
    """Try to read width from a door/window FamilyInstance."""
    return _get_opening_dimension(
        family_instance,
        (
            ("FAMILY_ROUGH_WIDTH_PARAM", "DOOR_ROUGH_WIDTH", "WINDOW_ROUGH_WIDTH"),
            ("FAMILY_WIDTH_PARAM", "DOOR_WIDTH", "WINDOW_WIDTH", "GENERIC_WIDTH"),
        ),
        (
            ("Rough Width", "Rough Opening Width"),
            ("Width",),
        ),
        3.0,
    )


def _get_opening_height(family_instance):
    """Try to read height from a door/window FamilyInstance."""
    return _get_opening_dimension(
        family_instance,
        (
            ("FAMILY_ROUGH_HEIGHT_PARAM", "DOOR_ROUGH_HEIGHT", "WINDOW_ROUGH_HEIGHT"),
            ("FAMILY_HEIGHT_PARAM", "DOOR_HEIGHT", "WINDOW_HEIGHT", "GENERIC_HEIGHT"),
        ),
        (
            ("Rough Height", "Rough Opening Height"),
            ("Height",),
        ),
        6.667,
    )


def _get_opening_dimension(family_instance, builtin_name_groups,
                           named_param_groups, default_value):
    """Read an opening dimension, preferring rough-opening parameters."""
    from Autodesk.Revit.DB import BuiltInParameter

    symbol = getattr(family_instance, "Symbol", None)
    for builtin_names, param_names in zip(builtin_name_groups, named_param_groups):
        for source in (family_instance, symbol):
            if source is None:
                continue
            for builtin_name in builtin_names:
                param_id = getattr(BuiltInParameter, builtin_name, None)
                if param_id is None:
                    continue
                try:
                    param = source.get_Parameter(param_id)
                except Exception:
                    param = None
                if param and param.HasValue:
                    value = param.AsDouble()
                    if value > 0.0:
                        return value
            for param_name in param_names:
                try:
                    param = source.LookupParameter(param_name)
                except Exception:
                    param = None
                if param and param.HasValue:
                    value = param.AsDouble()
                    if value > 0.0:
                        return value
    return default_value


def _get_sill_height(family_instance):
    """Try to read sill height of a window."""
    from Autodesk.Revit.DB import BuiltInParameter

    param_ids = []
    for name in (
        "INSTANCE_SILL_HEIGHT_PARAM",
        "FAMILY_SILL_HEIGHT_PARAM",
    ):
        pid = getattr(BuiltInParameter, name, None)
        if pid is not None:
            param_ids.append(pid)
    for pid in param_ids:
        try:
            p = family_instance.get_Parameter(pid)
        except Exception:
            p = None
        if p and p.HasValue:
            val = p.AsDouble()
            if val >= 0:
                return val

    for name in ("Sill Height",):
        try:
            p = family_instance.LookupParameter(name)
        except Exception:
            p = None
        if p and p.HasValue:
            val = p.AsDouble()
            if val >= 0:
                return val

    sym = getattr(family_instance, "Symbol", None)
    if sym is not None:
        for pid in param_ids:
            try:
                p = sym.get_Parameter(pid)
            except Exception:
                p = None
            if p and p.HasValue:
                val = p.AsDouble()
                if val >= 0:
                    return val
        for name in ("Sill Height",):
            try:
                p = sym.LookupParameter(name)
            except Exception:
                p = None
            if p and p.HasValue:
                val = p.AsDouble()
                if val >= 0:
                    return val
    return 3.0  # fallback 3 ft


def point_on_wall(wall_info, distance_along, height, lateral_offset=0.0):
    """Calculate a 3D point on the wall at a given distance and height.

    Args:
        wall_info: WallInfo
        distance_along: feet from wall start along wall direction
        height: feet above wall base (base_offset + level)
        lateral_offset: feet offset perpendicular to the wall.

    Returns:
        XYZ point
    """
    base_z = wall_info.level_elevation + wall_info.base_offset + height
    x = wall_info.start_point.X + wall_info.direction.X * distance_along
    y = wall_info.start_point.Y + wall_info.direction.Y * distance_along
    x += wall_info.normal.X * lateral_offset
    y += wall_info.normal.Y * lateral_offset
    return _make_xyz(x, y, base_z)


def safe_wall_normal(wall, direction=None):
    """Return a horizontal wall normal without raising on unsupported walls."""
    normal = None
    try:
        normal = wall.Orientation
    except Exception:
        normal = None

    if normal is not None:
        try:
            length = math.sqrt(normal.X * normal.X + normal.Y * normal.Y)
        except Exception:
            length = 0.0
        if length > 1e-9:
            return _make_xyz(normal.X / length, normal.Y / length, 0.0)

    if direction is not None:
        try:
            dx = direction.X
            dy = direction.Y
        except Exception:
            dx = 0.0
            dy = 0.0
        length = math.sqrt(dx * dx + dy * dy)
        if length > 1e-9:
            return _make_xyz(-dy / length, dx / length, 0.0)

    return None


def _make_xyz(x, y, z):
    """Create Revit XYZ point."""
    from Autodesk.Revit.DB import XYZ
    return XYZ(x, y, z)


# ------------------------------------------------------------------
# Slope detection
# ------------------------------------------------------------------

def height_at_position(wall_info, distance_along):
    """Get wall height at a position, interpolating for sloped walls.

    For flat-top walls, returns wall_info.height.
    For sloped walls, linearly interpolates between start and end heights.
    """
    if not wall_info.is_sloped_top:
        return wall_info.height
    if wall_info.length < 1e-9:
        return wall_info.start_height
    t = distance_along / wall_info.length
    t = max(0.0, min(1.0, t))
    return wall_info.start_height + (wall_info.end_height - wall_info.start_height) * t


def _detect_wall_slope(wall, info):
    """Detect if a wall has a sloped top (e.g. gable / vaulted).

    Examines the wall's solid geometry to find the highest Z at the
    wall start and end regions, comparing them to detect slope.
    """
    from Autodesk.Revit.DB import Options, Solid

    try:
        opt = Options()
        opt.ComputeReferences = False
        opt.IncludeNonVisibleObjects = False
        geom = wall.get_Geometry(opt)
        if geom is None:
            return
    except Exception:
        return

    # Collect all vertices from the wall solid
    all_pts = []
    for geom_obj in geom:
        solid = None
        if isinstance(geom_obj, Solid) and geom_obj.Volume > 0:
            solid = geom_obj
        if solid is None:
            continue
        for edge in solid.Edges:
            try:
                curve = edge.AsCurve()
                all_pts.append(curve.GetEndPoint(0))
                all_pts.append(curve.GetEndPoint(1))
            except Exception:
                pass

    if len(all_pts) < 4:
        return

    base_z = info.level_elevation + info.base_offset

    # Split points into "near start" and "near end" of wall
    # by projecting onto wall direction
    start_z_max = base_z
    end_z_max = base_z
    mid_dist = info.length / 2.0

    for pt in all_pts:
        dx = pt.X - info.start_point.X
        dy = pt.Y - info.start_point.Y
        along = dx * info.direction.X + dy * info.direction.Y

        if along <= mid_dist:
            if pt.Z > start_z_max:
                start_z_max = pt.Z
        else:
            if pt.Z > end_z_max:
                end_z_max = pt.Z

    # Convert absolute Z back to heights relative to wall base
    sh = start_z_max - base_z
    eh = end_z_max - base_z

    # Always update heights from actual geometry — WALL_USER_HEIGHT_PARAM
    # is "Unconnected Height" and wrong for walls with a Top Constraint.
    if sh > 0.0 and eh > 0.0:
        info.start_height = sh
        info.end_height = eh
        info.height = max(sh, eh)
        if abs(sh - eh) > inches_to_feet(2.0):
            info.is_sloped_top = True


def _set_wall_location_data(wall, info):
    """Capture the current wall location line and shift to core centerline."""
    from Autodesk.Revit.DB import WallLocationLine

    wall_type = wall.WallType
    if wall_type is None:
        return

    compound = wall_type.GetCompoundStructure()
    if compound is None:
        return

    current_line = _get_wall_location_line(wall)
    if current_line is None:
        current_line = WallLocationLine.WallCenterline

    try:
        current_offset = compound.GetOffsetForLocationLine(current_line)
        core_offset = compound.GetOffsetForLocationLine(
            WallLocationLine.CoreCenterline
        )
    except Exception:
        return

    info.location_line = current_line
    info.core_centerline_offset = core_offset - current_offset


def _get_wall_location_line(wall):
    """Return the wall's current location line enum value."""
    from Autodesk.Revit.DB import BuiltInParameter, WallLocationLine
    import System

    param = wall.get_Parameter(BuiltInParameter.WALL_KEY_REF_PARAM)
    if param is None or not param.HasValue:
        return None

    try:
        raw_value = param.AsInteger()
        if not System.Enum.IsDefined(WallLocationLine, raw_value):
            return None
        return System.Enum.ToObject(WallLocationLine, raw_value)
    except Exception:
        return None
