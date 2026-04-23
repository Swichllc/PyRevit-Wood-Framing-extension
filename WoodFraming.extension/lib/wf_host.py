# -*- coding: utf-8 -*-
"""Shared host and layer analysis for wall, floor, ceiling, and roof framing."""

import math

from wf_geometry import analyze_wall, find_openings, height_at_position
from wf_config import WALL_BASE_MODE_WALL, WALL_BASE_MODE_SUPPORT_TOP


LAYER_MODE_CORE_CENTER = "core_center"
LAYER_MODE_STRUCTURAL = "structural"
LAYER_MODE_THICKEST = "thickest"


class HostKind(object):
    """Supported host kinds."""

    WALL = "wall"
    FLOOR = "floor"
    CEILING = "ceiling"
    ROOF = "roof"


class HostLayerInfo(object):
    """Describes a single layer in a host compound structure."""

    def __init__(self):
        self.index = -1
        self.function = None
        self.material_id = None
        self.material_name = None
        self.width = 0.0
        self.start_depth = 0.0
        self.end_depth = 0.0
        self.depth_from_exterior = 0.0
        self.center_offset = 0.0
        self.is_core = False
        self.is_structural = False
        self.is_virtual = False


class HostElementInfo(object):
    """Base data shared by all analyzed hosts."""

    def __init__(self):
        self.kind = None
        self.element = None
        self.element_id = None
        self.level_id = None
        self.level_elevation = 0.0
        self.layers = []
        self.target_layer = None


class WallHostInfo(HostElementInfo):
    """Wall-specific host data in a stable local basis."""

    def __init__(self):
        HostElementInfo.__init__(self)
        self.wall_info = None
        self.start_point = None
        self.end_point = None
        self.direction = None
        self.normal = None
        self.up_axis = None
        self.length = 0.0
        self.height = 0.0
        self.start_height = 0.0
        self.end_height = 0.0
        self.is_sloped_top = False
        self.base_offset = 0.0
        self.base_elevation = 0.0
        self.location_line = None
        self.current_location_line_offset = 0.0
        self.target_layer_offset = 0.0
        self.openings = []

    def point_at(self, distance_along, height, lateral_offset=0.0):
        """Return a world-space point on the selected wall framing line."""
        from Autodesk.Revit.DB import XYZ

        base_z = self.base_elevation + height
        x = self.start_point.X + self.direction.X * distance_along
        y = self.start_point.Y + self.direction.Y * distance_along
        offset = self.target_layer_offset + lateral_offset
        x += self.normal.X * offset
        y += self.normal.Y * offset
        return XYZ(x, y, base_z)

    def height_at(self, distance_along):
        """Return the actual wall height at a distance along the wall."""
        return height_at_position(self.wall_info, distance_along)


class PlanarHostInfo(HostElementInfo):
    """Planar host face data for floor, ceiling, or roof framing."""

    def __init__(self):
        HostElementInfo.__init__(self)
        self.face = None
        self.face_index = 0
        self.origin = None
        self.x_axis = None
        self.y_axis = None
        self.normal = None
        self.boundary_loops_local = []
        self.outer_loop_local = []
        self.bounds = (0.0, 0.0, 0.0, 0.0)
        self.area = 0.0
        self.target_layer_depth = 0.0

    def point_at(self, local_x, local_y, depth_offset=0.0):
        """Convert local host coordinates into a world-space point."""
        total_depth = self.target_layer_depth + depth_offset
        return _point_from_axes(
            self.origin,
            self.x_axis,
            self.y_axis,
            self.normal,
            local_x,
            local_y,
            -total_depth,
        )

    def scanline_intervals(self, axis_name, coord):
        """Return inside intervals for a scanline through the host profile."""
        return _scanline_intervals(self.boundary_loops_local, axis_name, coord)


class RoofHostInfo(HostElementInfo):
    """Roof host data containing one planar framing plane per roof face."""

    def __init__(self):
        HostElementInfo.__init__(self)
        self.planes = []


def analyze_wall_host(doc, wall, config):
    """Analyze a wall into a host-local framing model."""
    wall_info = analyze_wall(doc, wall)
    if wall_info is None:
        return None

    compound = _get_compound_structure(wall)
    layers = _build_compound_layers(doc, compound)
    mode = getattr(config, "wall_layer_mode", LAYER_MODE_CORE_CENTER)
    target_layer = _select_target_layer(layers, mode)
    target_layer = _preferred_wall_target_layer(layers, mode, target_layer)

    # Shift from Location.Curve (at wall center) to core centerline.
    # GetOffsetForLocationLine(CoreCenterline) returns the offset from the
    # compound structure center (= wall center) to core center.
    # Apply this directly — do NOT subtract the current location line offset.
    core_shift = 0.0
    if compound is not None:
        from Autodesk.Revit.DB import WallLocationLine
        try:
            core_shift = compound.GetOffsetForLocationLine(
                WallLocationLine.CoreCenterline)
        except Exception:
            core_shift = 0.0

    info = WallHostInfo()
    info.kind = HostKind.WALL
    info.element = wall
    info.element_id = wall.Id
    info.level_id = wall_info.level_id
    info.level_elevation = wall_info.level_elevation
    info.wall_info = wall_info
    info.start_point = wall_info.start_point
    info.end_point = wall_info.end_point
    info.direction = wall_info.direction
    info.normal = wall_info.normal
    info.up_axis = _world_up()
    info.length = wall_info.length
    info.height = wall_info.height
    info.start_height = wall_info.start_height
    info.end_height = wall_info.end_height
    info.is_sloped_top = wall_info.is_sloped_top
    info.base_offset = wall_info.base_offset
    info.base_elevation = wall_info.level_elevation + wall_info.base_offset
    info.location_line = wall_info.location_line
    info.current_location_line_offset = 0.0
    info.layers = layers
    info.target_layer = target_layer
    info.target_layer_offset = core_shift
    info.openings = find_openings(doc, wall, wall_info)

    # Optional override: frame wall from selected support-host top elevation.
    try:
        apply_wall_base_override_from_config(doc, info, config)
    except Exception:
        pass

    return info


def apply_wall_base_override_from_config(doc, wall_host_info, config):
    """Shift wall framing baseline to selected support-host top elevation."""
    if wall_host_info is None or config is None:
        return False

    mode = getattr(config, "wall_base_mode", WALL_BASE_MODE_WALL)
    if mode != WALL_BASE_MODE_SUPPORT_TOP:
        return False

    support_id = getattr(config, "wall_base_support_element_id", None)
    support = _get_element_by_any_id(doc, support_id)
    if support is None:
        return False

    target_base = _resolve_support_top_for_wall(doc, support, wall_host_info)
    if target_base is None:
        return False

    return _shift_wall_host_base(wall_host_info, target_base)


def _get_element_by_any_id(doc, raw_id):
    """Resolve an element from an int, ElementId, or string id."""
    if raw_id is None:
        return None

    try:
        # Already an ElementId-like object.
        if hasattr(raw_id, "IntegerValue") or hasattr(raw_id, "Value"):
            return doc.GetElement(raw_id)
    except Exception:
        pass

    try:
        from Autodesk.Revit.DB import ElementId

        return doc.GetElement(ElementId(int(raw_id)))
    except Exception:
        return None


def _resolve_support_top_for_wall(doc, support, wall_host_info):
    """Sample support top elevation at key wall stations and return a base Z."""
    length = max(0.0, getattr(wall_host_info, "length", 0.0))
    station_count = 8
    stations = []
    for index in range(station_count + 1):
        stations.append(length * (float(index) / float(station_count)))

    lateral_offsets = [0.0]
    wall_info = getattr(wall_host_info, "wall_info", None)
    wall_width = float(getattr(wall_info, "width", 0.0) or 0.0)
    target_layer = getattr(wall_host_info, "target_layer", None)
    target_width = float(getattr(target_layer, "width", 0.0) or 0.0)
    half_span = max(wall_width * 0.5, target_width * 0.5)
    if half_span > 1e-6:
        lateral_offsets.extend([half_span, -half_span, half_span * 0.5, -half_span * 0.5])

    tops = []
    for station in stations:
        for lateral in lateral_offsets:
            try:
                sample = wall_host_info.point_at(station, 0.0, lateral)
            except Exception:
                continue
            top_z = _solid_top_elevation_at_xy(support, sample)
            if top_z is not None:
                tops.append(top_z)

        # Fallback sample on the raw wall location line (no layer offset).
        try:
            raw_x = wall_host_info.start_point.X + wall_host_info.direction.X * station
            raw_y = wall_host_info.start_point.Y + wall_host_info.direction.Y * station
            from Autodesk.Revit.DB import XYZ

            raw_sample = XYZ(raw_x, raw_y, wall_host_info.base_elevation)
            raw_top_z = _solid_top_elevation_at_xy(support, raw_sample)
            if raw_top_z is not None:
                tops.append(raw_top_z)
        except Exception:
            pass

    if tops:
        return min(tops)

    # If ray sampling failed, still return a usable baseline estimate.
    try:
        bbox = support.get_BoundingBox(None)
        if bbox is not None:
            return bbox.Max.Z
    except Exception:
        pass

    return None


def _solid_top_elevation_at_xy(element, sample_point):
    """Return top Z where a vertical ray at XY intersects element solids."""
    try:
        from Autodesk.Revit.DB import (
            GeometryInstance,
            Line,
            Options,
            Solid,
            SolidCurveIntersectionOptions,
            ViewDetailLevel,
            XYZ,
        )
    except Exception:
        return None

    try:
        opts = Options()
        opts.ComputeReferences = False
        opts.DetailLevel = ViewDetailLevel.Fine
        geom = element.get_Geometry(opts)
    except Exception:
        geom = None
    if geom is None:
        return None

    solids = []
    for geom_obj in geom:
        if isinstance(geom_obj, Solid) and geom_obj.Volume > 0:
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
                if isinstance(sub, Solid) and sub.Volume > 0:
                    solids.append(sub)

    if not solids:
        return None

    try:
        bbox = element.get_BoundingBox(None)
    except Exception:
        bbox = None
    if bbox is None:
        return None

    z_min = bbox.Min.Z - 10.0
    z_max = bbox.Max.Z + 10.0
    if z_max - z_min < 1e-6:
        return None

    try:
        ray = Line.CreateBound(
            XYZ(sample_point.X, sample_point.Y, z_min),
            XYZ(sample_point.X, sample_point.Y, z_max),
        )
    except Exception:
        return None

    top_z = None
    for solid in solids:
        try:
            result = solid.IntersectWithCurve(
                ray,
                SolidCurveIntersectionOptions(),
            )
            seg_count = result.SegmentCount
        except Exception:
            continue

        for index in range(seg_count):
            try:
                segment = result.GetCurveSegment(index)
                z0 = segment.GetEndPoint(0).Z
                z1 = segment.GetEndPoint(1).Z
            except Exception:
                continue
            seg_top = max(z0, z1)
            if top_z is None or seg_top > top_z:
                top_z = seg_top

    return top_z


def _shift_wall_host_base(wall_host_info, new_base_elevation):
    """Shift wall-host local Z baseline while preserving world-space top geometry."""
    old_base = getattr(wall_host_info, "base_elevation", None)
    if old_base is None:
        return False

    delta = new_base_elevation - old_base
    if abs(delta) < 1e-9:
        return False

    wall_host_info.base_elevation = new_base_elevation
    wall_host_info.base_offset = new_base_elevation - wall_host_info.level_elevation

    wall_host_info.height = max(0.0, wall_host_info.height - delta)
    wall_host_info.start_height = max(0.0, wall_host_info.start_height - delta)
    wall_host_info.end_height = max(0.0, wall_host_info.end_height - delta)
    wall_host_info.is_sloped_top = (
        abs(wall_host_info.start_height - wall_host_info.end_height) > 1e-6
    )

    wall_info = getattr(wall_host_info, "wall_info", None)
    if wall_info is not None:
        wall_info.base_offset = wall_host_info.base_offset
        wall_info.height = wall_host_info.height
        wall_info.start_height = wall_host_info.start_height
        wall_info.end_height = wall_host_info.end_height
        wall_info.is_sloped_top = wall_host_info.is_sloped_top

    for opening in getattr(wall_host_info, "openings", []):
        try:
            opening.sill_height -= delta
            opening.head_height -= delta
        except Exception:
            continue

    return True


def analyze_floor_host(doc, floor, config):
    """Analyze a floor using the shared planar-host pipeline."""
    layer_mode = getattr(config, "floor_layer_mode", LAYER_MODE_STRUCTURAL)
    return _analyze_single_planar_host(doc, floor, HostKind.FLOOR, layer_mode)


def analyze_ceiling_host(doc, ceiling, config):
    """Analyze a ceiling using the shared planar-host pipeline."""
    layer_mode = getattr(config, "ceiling_layer_mode", LAYER_MODE_STRUCTURAL)
    return _analyze_single_planar_host(doc, ceiling, HostKind.CEILING, layer_mode)


def analyze_roof_host(doc, roof, config):
    """Analyze a roof into one framing plane per top face."""
    roof_info = RoofHostInfo()
    roof_info.kind = HostKind.ROOF
    roof_info.element = roof
    roof_info.element_id = roof.Id
    roof_info.level_id = roof.LevelId
    roof_info.level_elevation = _get_level_elevation(doc, roof.LevelId)

    compound = _get_compound_structure(roof)
    layers = _build_compound_layers(doc, compound)
    layer_mode = getattr(config, "roof_layer_mode", LAYER_MODE_STRUCTURAL)
    target_layer = _select_target_layer(layers, layer_mode)

    faces = _get_top_faces(roof)
    for index, face in enumerate(faces):
        plane = _build_planar_host_info(
            doc,
            roof,
            HostKind.ROOF,
            face,
            layers,
            target_layer,
            index,
            roof_mode=True,
        )
        if plane is not None:
            roof_info.planes.append(plane)

    roof_info.planes.sort(key=lambda plane: plane.area, reverse=True)
    roof_info.layers = layers
    roof_info.target_layer = target_layer
    if not roof_info.planes:
        return None
    return roof_info


def _analyze_single_planar_host(doc, host, kind, layer_mode):
    """Analyze a floor-like or ceiling-like host with a single dominant face."""
    faces = _get_horizontal_faces(host)
    if not faces:
        return None

    compound = _get_compound_structure(host)
    layers = _build_compound_layers(doc, compound)
    target_layer = _select_target_layer(layers, layer_mode)

    # find largest face
    best_area = -1.0
    largest_faces = []
    for face in faces:
        area = _get_face_area(face)
        if area > best_area + 0.1:
            best_area = area
            largest_faces = [face]
        elif abs(area - best_area) <= 0.1:
            largest_faces.append(face)

    # From largest faces, pick the one with highest Z elevation
    from Autodesk.Revit.DB import UV
    best_face = None
    highest_z = -1e9
    for face in largest_faces:
        try:
            bbox = face.GetBoundingBox()
            center_uv = UV(
                (bbox.Min.U + bbox.Max.U) / 2.0,
                (bbox.Min.V + bbox.Max.V) / 2.0
            )
            pt = face.Evaluate(center_uv)
            if pt.Z > highest_z:
                highest_z = pt.Z
                best_face = face
        except Exception:
            pass

    if best_face is None:
        return None

    return _build_planar_host_info(
        doc,
        host,
        kind,
        best_face,
        layers,
        target_layer,
        0,
        roof_mode=False,
    )


def _build_planar_host_info(doc, host, kind, face, layers,
                            target_layer, face_index, roof_mode=False):
    """Build a planar host model from a single top face."""
    normal = _face_normal(face)
    if normal is None:
        return None
    if normal.Z < 0.0:
        normal = normal.Multiply(-1.0)

    loops = _extract_face_loops(face)
    if not loops:
        return None

    origin = loops[0][0]
    x_axis, y_axis = _choose_planar_axes(loops, normal, roof_mode)
    if x_axis is None or y_axis is None:
        return None

    loops_local = []
    for loop in loops:
        local_loop = []
        for point in loop:
            local_loop.append(_to_local(point, origin, x_axis, y_axis))
        if len(local_loop) >= 3:
            loops_local.append(local_loop)

    if not loops_local:
        return None

    outer_loop_local = _pick_outer_loop(loops_local)
    if not outer_loop_local:
        return None

    min_x, max_x, min_y, max_y = _loop_bounds(outer_loop_local)

    info = PlanarHostInfo()
    info.kind = kind
    info.element = host
    info.element_id = host.Id
    info.level_id = host.LevelId
    info.level_elevation = _get_level_elevation(doc, host.LevelId)
    info.layers = layers
    info.target_layer = target_layer
    info.face = face
    info.face_index = face_index
    info.origin = origin
    info.x_axis = x_axis
    info.y_axis = y_axis
    info.normal = normal
    info.boundary_loops_local = loops_local
    info.outer_loop_local = outer_loop_local
    info.bounds = (min_x, max_x, min_y, max_y)
    info.area = abs(_polygon_area_2d(outer_loop_local))
    if target_layer is not None:
        info.target_layer_depth = target_layer.depth_from_exterior
    return info


def _get_compound_structure(host):
    """Get the host type compound structure when available."""
    try:
        type_elem = host.Document.GetElement(host.GetTypeId())
    except Exception:
        type_elem = None
    if type_elem is None:
        return None
    get_structure = getattr(type_elem, "GetCompoundStructure", None)
    if get_structure is None:
        return None
    try:
        return get_structure()
    except Exception:
        return None


def _build_compound_layers(doc, compound):
    """Translate a Revit CompoundStructure into framing layer data."""
    if compound is None:
        return []

    try:
        source_layers = list(compound.GetLayers())
    except Exception:
        return []

    total_width = 0.0
    for source_layer in source_layers:
        try:
            total_width += max(0.0, source_layer.Width)
        except Exception:
            pass

    try:
        first_core = compound.GetFirstCoreLayerIndex()
        last_core = compound.GetLastCoreLayerIndex()
    except Exception:
        first_core = -1
        last_core = -1

    try:
        structural_index = compound.StructuralMaterialIndex
    except Exception:
        structural_index = -1

    layers = []
    depth_cursor = 0.0
    for index, source_layer in enumerate(source_layers):
        width = 0.0
        try:
            width = max(0.0, source_layer.Width)
        except Exception:
            width = 0.0

        layer = HostLayerInfo()
        layer.index = index
        layer.width = width
        layer.start_depth = depth_cursor
        layer.end_depth = depth_cursor + width
        layer.depth_from_exterior = depth_cursor + (width / 2.0)
        layer.center_offset = (total_width / 2.0) - layer.depth_from_exterior
        layer.function = _enum_name(getattr(source_layer, "Function", None))
        layer.material_id = getattr(source_layer, "MaterialId", None)
        layer.material_name = _material_name(doc, layer.material_id)
        layer.is_core = (first_core >= 0 and first_core <= index <= last_core)
        layer.is_structural = (structural_index == index)
        layers.append(layer)
        depth_cursor += width

    return layers


def _select_target_layer(layers, mode):
    """Select the layer used as the framing control line or plane."""
    positive_layers = [layer for layer in layers if layer.width > 1e-9]
    if not positive_layers:
        return None

    if mode == LAYER_MODE_CORE_CENTER:
        core_center = _make_core_center_layer(positive_layers)
        if core_center is not None:
            return core_center

    if mode == LAYER_MODE_STRUCTURAL:
        for layer in positive_layers:
            if layer.is_structural:
                return layer
        core_center = _make_core_center_layer(positive_layers)
        if core_center is not None:
            return core_center

    if mode == LAYER_MODE_THICKEST:
        return max(positive_layers, key=lambda layer: layer.width)

    core_layers = [layer for layer in positive_layers if layer.is_core]
    if core_layers:
        return max(core_layers, key=lambda layer: layer.width)
    return max(positive_layers, key=lambda layer: layer.width)


def _preferred_wall_target_layer(layers, mode, target_layer):
    """Prefer the actual structural layer over a virtual core centerline."""
    if mode != LAYER_MODE_CORE_CENTER or target_layer is None:
        return target_layer
    if not getattr(target_layer, "is_virtual", False):
        return target_layer

    for layer in layers:
        if layer.width > 1e-9 and layer.is_structural:
            return layer
    return target_layer


def _make_core_center_layer(layers):
    """Create a virtual layer representing the entire host core."""
    core_layers = [layer for layer in layers if layer.is_core and layer.width > 1e-9]
    if not core_layers:
        return None

    layer = HostLayerInfo()
    layer.index = core_layers[0].index
    layer.width = core_layers[-1].end_depth - core_layers[0].start_depth
    layer.start_depth = core_layers[0].start_depth
    layer.end_depth = core_layers[-1].end_depth
    layer.depth_from_exterior = layer.start_depth + (layer.width / 2.0)
    total_width = layers[-1].end_depth if layers else 0.0
    layer.center_offset = (total_width / 2.0) - layer.depth_from_exterior
    layer.is_core = True
    layer.is_structural = any(item.is_structural for item in core_layers)
    layer.is_virtual = True
    layer.function = "CoreCenter"
    return layer


def _get_top_faces(host):
    """Get all major top faces on a host object."""
    from Autodesk.Revit.DB import HostObjectUtils, Options

    faces = []
    try:
        refs = HostObjectUtils.GetTopFaces(host)
    except Exception:
        refs = []

    for reference in refs:
        try:
            face = host.GetGeometryObjectFromReference(reference)
        except Exception:
            face = None
        if face is not None:
            normal = _face_normal(face)
            if normal is not None and normal.Z > 0.01:
                faces.append(face)

    if faces:
        return faces

    # Fallback for hosts where GetTopFaces is unavailable.
    try:
        opt = Options()
        opt.ComputeReferences = True
        geom = host.get_Geometry(opt)
    except Exception:
        geom = None

    if geom is None:
        return faces

    for geom_obj in geom:
        solids = []
        if hasattr(geom_obj, "Faces"):
            solids.append(geom_obj)
        elif hasattr(geom_obj, "GetInstanceGeometry"):
            try:
                inst_geom = geom_obj.GetInstanceGeometry()
            except Exception:
                inst_geom = None
            if inst_geom:
                for inst_obj in inst_geom:
                    if hasattr(inst_obj, "Faces"):
                        solids.append(inst_obj)

        for solid in solids:
            for face in solid.Faces:
                normal = _face_normal(face)
                if normal is not None and normal.Z > 0.01:
                    faces.append(face)

    return faces


def _get_horizontal_faces(host):
    """Get all major horizontal faces on a host object (top and bottom)."""
    from Autodesk.Revit.DB import HostObjectUtils, Options

    faces = []
    refs = []
    try:
        t_refs = HostObjectUtils.GetTopFaces(host)
        for r in t_refs: refs.append(r)
    except Exception:
        pass
        
    try:
        b_refs = HostObjectUtils.GetBottomFaces(host)
        for r in b_refs: refs.append(r)
    except Exception:
        pass

    for reference in refs:
        try:
            face = host.GetGeometryObjectFromReference(reference)
        except Exception:
            face = None
        if face is not None:
            normal = _face_normal(face)
            if normal is not None and abs(normal.Z) > 0.01:
                faces.append(face)

    if faces:
        return faces

    # Fallback for hosts where GetTopFaces/GetBottomFaces is unavailable.
    try:
        opt = Options()
        opt.ComputeReferences = True
        geom = host.get_Geometry(opt)
    except Exception:
        geom = None

    if geom is None:
        return faces

    for geom_obj in geom:
        solids = []
        if hasattr(geom_obj, "Faces"):
            solids.append(geom_obj)
        elif hasattr(geom_obj, "GetInstanceGeometry"):
            try:
                inst_geom = geom_obj.GetInstanceGeometry()
            except Exception:
                inst_geom = None
            if inst_geom:
                for inst_obj in inst_geom:
                    if hasattr(inst_obj, "Faces"):
                        solids.append(inst_obj)

        for solid in solids:
            for face in solid.Faces:
                normal = _face_normal(face)
                if normal is not None and abs(normal.Z) > 0.01:
                    faces.append(face)

    return faces


def _extract_face_loops(face):
    """Extract tessellated loop points from a face."""
    loops = []
    try:
        curve_loops = face.GetEdgesAsCurveLoops()
    except Exception:
        return loops

    for curve_loop in curve_loops:
        points = []
        for curve in curve_loop:
            try:
                tessellated = list(curve.Tessellate())
            except Exception:
                tessellated = [curve.GetEndPoint(0), curve.GetEndPoint(1)]
            for point in tessellated:
                if points and _points_close(points[-1], point):
                    continue
                points.append(point)

        if len(points) >= 2 and _points_close(points[0], points[-1]):
            points = points[:-1]
        if len(points) >= 3:
            loops.append(points)

    return loops


def _choose_planar_axes(loops, normal, roof_mode):
    """Choose stable in-plane axes for floor or roof framing."""
    if roof_mode:
        y_axis = _downhill_direction(normal)
        if y_axis is None:
            y_axis = _longest_edge_direction(loops, normal)
        if y_axis is None:
            return None, None
        x_axis = _normalize(normal.CrossProduct(y_axis))
    else:
        x_axis = _longest_edge_direction(loops, normal)
        if x_axis is None:
            x_axis = _project_to_plane(_world_x(), normal)
            x_axis = _normalize(x_axis)
        if x_axis is None:
            return None, None
        y_axis = _normalize(normal.CrossProduct(x_axis))

    if x_axis is None or y_axis is None:
        return None, None
    return x_axis, y_axis


def _downhill_direction(normal):
    """Return the steepest in-plane downward direction for a roof face."""
    downhill = _project_to_plane(_world_up().Multiply(-1.0), normal)
    return _normalize(downhill)


def _longest_edge_direction(loops, normal):
    """Return the longest useful in-plane edge direction from the loops."""
    best_length = 0.0
    best_direction = None
    for loop in loops:
        count = len(loop)
        for index in range(count):
            start = loop[index]
            end = loop[(index + 1) % count]
            vector = end - start
            vector = _project_to_plane(vector, normal)
            length = _vector_length(vector)
            if length > best_length and length > 1e-9:
                best_length = length
                best_direction = _normalize(vector)
    return best_direction


def _pick_outer_loop(loops_local):
    """Pick the largest loop by absolute projected area."""
    best_loop = None
    best_area = 0.0
    for loop in loops_local:
        area = abs(_polygon_area_2d(loop))
        if area > best_area:
            best_area = area
            best_loop = loop
    return best_loop


def _loop_bounds(loop):
    """Return min/max bounds for a local loop."""
    min_x = min(point[0] for point in loop)
    max_x = max(point[0] for point in loop)
    min_y = min(point[1] for point in loop)
    max_y = max(point[1] for point in loop)
    return min_x, max_x, min_y, max_y


def _scanline_intervals(loops_local, axis_name, coord):
    """Compute inside intervals for a 2D scanline using even-odd pairing."""
    values = []
    for loop in loops_local:
        count = len(loop)
        for index in range(count):
            start = loop[index]
            end = loop[(index + 1) % count]
            if axis_name == "y":
                intersection = _scan_y_intersection(start, end, coord)
            else:
                intersection = _scan_x_intersection(start, end, coord)
            if intersection is not None:
                values.append(intersection)

    values.sort()
    intervals = []
    pair_index = 0
    while pair_index + 1 < len(values):
        start_value = values[pair_index]
        end_value = values[pair_index + 1]
        if end_value - start_value > 1e-6:
            intervals.append((start_value, end_value))
        pair_index += 2
    return intervals


def _scan_y_intersection(start, end, scan_y):
    """Intersect a segment with a constant-Y scanline."""
    y1 = start[1]
    y2 = end[1]
    if abs(y1 - y2) < 1e-9:
        return None
    lower = min(y1, y2)
    upper = max(y1, y2)
    if scan_y < lower or scan_y >= upper:
        return None
    t = (scan_y - y1) / (y2 - y1)
    return start[0] + (end[0] - start[0]) * t


def _scan_x_intersection(start, end, scan_x):
    """Intersect a segment with a constant-X scanline."""
    x1 = start[0]
    x2 = end[0]
    if abs(x1 - x2) < 1e-9:
        return None
    lower = min(x1, x2)
    upper = max(x1, x2)
    if scan_x < lower or scan_x >= upper:
        return None
    t = (scan_x - x1) / (x2 - x1)
    return start[1] + (end[1] - start[1]) * t


def _face_normal(face):
    """Get a representative face normal."""
    from Autodesk.Revit.DB import UV

    try:
        bbox = face.GetBoundingBox()
        uv = UV(
            (bbox.Min.U + bbox.Max.U) / 2.0,
            (bbox.Min.V + bbox.Max.V) / 2.0,
        )
        return _normalize(face.ComputeNormal(uv))
    except Exception:
        return None


def _get_face_area(face):
    """Return the face area when available."""
    try:
        return face.Area
    except Exception:
        return 0.0


def _to_local(point, origin, x_axis, y_axis):
    """Project a world-space point into a planar host basis."""
    vector = point - origin
    return (vector.DotProduct(x_axis), vector.DotProduct(y_axis))


def _point_from_axes(origin, x_axis, y_axis, z_axis, local_x, local_y, local_z):
    """Compose a world-space point from a local basis."""
    return (
        origin
        + x_axis.Multiply(local_x)
        + y_axis.Multiply(local_y)
        + z_axis.Multiply(local_z)
    )


def _vector_length(vector):
    """Return vector length."""
    try:
        return vector.GetLength()
    except Exception:
        return 0.0


def _normalize(vector):
    """Return a unit vector or None."""
    if vector is None:
        return None
    try:
        if vector.GetLength() < 1e-9:
            return None
        return vector.Normalize()
    except Exception:
        return None


def _project_to_plane(vector, normal):
    """Project a vector onto the plane orthogonal to normal."""
    try:
        return vector - normal.Multiply(vector.DotProduct(normal))
    except Exception:
        return None


def _polygon_area_2d(points):
    """Return the signed area of a 2D polygon."""
    area = 0.0
    count = len(points)
    for index in range(count):
        x1, y1 = points[index]
        x2, y2 = points[(index + 1) % count]
        area += (x1 * y2) - (x2 * y1)
    return area / 2.0


def _points_close(first, second):
    """Check whether two XYZ points are effectively identical."""
    try:
        return first.DistanceTo(second) < 1e-6
    except Exception:
        return False


def _world_up():
    """Return the global up vector."""
    from Autodesk.Revit.DB import XYZ

    return XYZ(0.0, 0.0, 1.0)


def _world_x():
    """Return the global X axis."""
    from Autodesk.Revit.DB import XYZ

    return XYZ(1.0, 0.0, 0.0)


def _get_level_elevation(doc, level_id):
    """Return level elevation for a host."""
    level = doc.GetElement(level_id)
    if level is None:
        return 0.0
    try:
        return level.Elevation
    except Exception:
        return 0.0


def _material_name(doc, material_id):
    """Return the material name for a layer."""
    if material_id is None:
        return None
    try:
        value = getattr(material_id, "IntegerValue", getattr(material_id, "Value", -1))
        if value < 0:
            return None
    except Exception:
        return None
    material = doc.GetElement(material_id)
    if material is None:
        return None
    try:
        return material.Name
    except Exception:
        return None


def _enum_name(value):
    """Return a stable enum-like string."""
    if value is None:
        return None
    try:
        return value.ToString()
    except Exception:
        return str(value)