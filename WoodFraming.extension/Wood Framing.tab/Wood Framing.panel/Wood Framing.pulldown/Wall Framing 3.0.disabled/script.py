# -*- coding: utf-8 -*-
"""Wall Framing 3.0 - standalone wall framing command.

This file is intentionally self-contained. It does not import the existing
wood-framing project libraries and it does not delete or update existing
generated framing.
"""

import math

from pyrevit import DB, forms, revit, script
from Autodesk.Revit.UI.Selection import ISelectionFilter, ObjectType
from Autodesk.Revit.DB.Structure import StructuralType


ENGINE_NAME = "wall-framing-3.0-standalone"
TRACK_PREFIX = "WF3|"
MIN_LEN = 1.0 / 12.0
PLATE_T = 1.5 / 12.0
STUD_T = 1.5 / 12.0
DEFAULT_HEADER_DEPTH = 3.5 / 12.0


class WallSelectionFilter(ISelectionFilter):
    def AllowElement(self, element):
        return isinstance(element, DB.Wall)

    def AllowReference(self, reference, point):
        return False


class Opening3(object):
    def __init__(self, left, right, sill, head, is_window):
        self.left = max(0.0, left)
        self.right = max(self.left, right)
        self.sill = max(0.0, sill)
        self.head = max(self.sill, head)
        self.is_window = bool(is_window)


class Host3(object):
    def __init__(self):
        self.wall = None
        self.wall_id = None
        self.level = None
        self.base_z = 0.0
        self.start = None
        self.direction = None
        self.interior_normal = None
        self.length = 0.0
        self.outline_2d = []
        self.openings = []
        self.perimeter_segments = []
        self.audit = {}


class Member3(object):
    def __init__(self, kind, start, end, symbol, is_column, rotation=0.0):
        self.kind = kind
        self.start = start
        self.end = end
        self.symbol = symbol
        self.is_column = bool(is_column)
        self.rotation = rotation


class FaceSegment3(object):
    def __init__(self, kind, p0, p1, d0, d1):
        self.kind = kind
        self.p0 = p0
        self.p1 = p1
        self.d0 = d0
        self.d1 = d1


def main():
    doc = revit.doc
    walls = selected_or_picked_walls(doc)
    if not walls:
        return

    stud_symbol = choose_symbol(
        doc,
        DB.BuiltInCategory.OST_StructuralColumns,
        "Select stud column type",
    )
    if stud_symbol is None:
        return
    plate_symbol = choose_symbol(
        doc,
        DB.BuiltInCategory.OST_StructuralFraming,
        "Select plate framing type",
    )
    if plate_symbol is None:
        return
    header_symbol = choose_symbol(
        doc,
        DB.BuiltInCategory.OST_StructuralFraming,
        "Select header framing type",
    )
    if header_symbol is None:
        return

    spacing_in = ask_positive_float("Stud spacing in inches", "16")
    if spacing_in is None:
        return
    top_count = ask_positive_int("Top plate count", "2")
    if top_count is None:
        return

    options = {
        "stud": stud_symbol,
        "plate": plate_symbol,
        "header": header_symbol,
        "spacing": spacing_in / 12.0,
        "top_count": top_count,
        "bottom_count": 1,
    }

    output = script.get_output()
    audit_rows = []
    placed_total = 0
    skipped = 0

    with revit.Transaction("WF: Wall Framing 3.0 Standalone"):
        activate_symbol(doc, stud_symbol)
        activate_symbol(doc, plate_symbol)
        activate_symbol(doc, header_symbol)
        doc.Regenerate()

        for wall in walls:
            host = analyze_wall_3(doc, wall)
            if host is None:
                skipped += 1
                continue
            members = build_members_3(host, options)
            placed = place_members_3(doc, host, members)
            placed_total += placed
            audit_rows.append((host.audit, placed))

    output.print_md(build_report(len(walls), skipped, placed_total, audit_rows))


def selected_or_picked_walls(doc):
    selected = revit.get_selection().elements
    walls = [element for element in selected if isinstance(element, DB.Wall)]
    if walls:
        return walls
    try:
        refs = revit.uidoc.Selection.PickObjects(
            ObjectType.Element,
            WallSelectionFilter(),
            "Select walls for Wall Framing 3.0",
        )
    except Exception:
        return []
    result = []
    for ref in refs:
        wall = doc.GetElement(ref.ElementId)
        if isinstance(wall, DB.Wall):
            result.append(wall)
    return result


def collect_symbols(doc, category):
    collector = (
        DB.FilteredElementCollector(doc)
        .OfCategory(category)
        .OfClass(DB.FamilySymbol)
    )
    items = []
    for symbol in collector:
        label = symbol_label(symbol)
        items.append((label, symbol))
    items.sort(key=lambda item: item[0].lower())
    return items


def choose_symbol(doc, category, title):
    items = collect_symbols(doc, category)
    if not items:
        forms.alert("No family types found for: {0}".format(title), title="Wall Framing 3.0")
        return None
    labels = [label for label, _ in items]
    selected = forms.SelectFromList.show(labels, title=title, multiselect=False)
    if not selected:
        return None
    for label, symbol in items:
        if label == selected:
            return symbol
    return None


def ask_positive_float(prompt, default):
    value = forms.ask_for_string(default=default, prompt=prompt, title="Wall Framing 3.0")
    if value is None:
        return None
    try:
        parsed = float(value)
    except Exception:
        parsed = 0.0
    if parsed <= 0.0:
        forms.alert("Value must be greater than zero.", title="Wall Framing 3.0")
        return None
    return parsed


def ask_positive_int(prompt, default):
    value = forms.ask_for_string(default=default, prompt=prompt, title="Wall Framing 3.0")
    if value is None:
        return None
    try:
        parsed = int(value)
    except Exception:
        parsed = 0
    if parsed <= 0:
        forms.alert("Value must be greater than zero.", title="Wall Framing 3.0")
        return None
    return parsed


def analyze_wall_3(doc, wall):
    loc = getattr(wall, "Location", None)
    curve = getattr(loc, "Curve", None)
    if curve is None or not isinstance(curve, DB.Line):
        return None

    p0 = curve.GetEndPoint(0)
    p1 = curve.GetEndPoint(1)
    direction = normalize_xyz(p1 - p0)
    if direction is None:
        return None

    level = doc.GetElement(wall.LevelId)
    if level is None:
        return None
    base_z = level.Elevation + param_double(wall, DB.BuiltInParameter.WALL_BASE_OFFSET, 0.0)

    face = largest_interior_face(wall)
    if face is None:
        return None

    loops3 = face_curve_loops(face)
    if not loops3:
        return None

    interior_normal = interior_face_normal(face, wall)
    if interior_normal is None:
        return None

    anchor = first_point(loops3)
    if anchor is None:
        return None

    p0_on_face = p0 - interior_normal * ((p0 - anchor).DotProduct(interior_normal))
    depth_from_interior, layer_label = target_depth_from_interior(doc, wall)
    target_origin = p0_on_face - interior_normal * depth_from_interior
    target_origin = DB.XYZ(target_origin.X, target_origin.Y, base_z)

    loops2 = loops_local_2d(loops3, p0_on_face, direction)
    outer_index = largest_loop_index(loops2)
    if outer_index is None:
        return None
    outer = loops2[outer_index]
    start_d = min(d for d, _ in outer)
    end_d = max(d for d, _ in outer)
    length = end_d - start_d
    if length < MIN_LEN:
        return None
    outer_local = [(d - start_d, z) for d, z in outer]

    face_openings = []
    for index, loop in enumerate(loops2):
        if index == outer_index:
            continue
        opening = opening_from_loop(loop, start_d, length, base_z)
        if opening is not None:
            face_openings.append(opening)

    hosted_openings = hosted_openings_3(doc, wall, p0_on_face, direction, start_d, end_d, base_z)
    openings = merge_openings(hosted_openings, face_openings, length)
    perimeter_segments = perimeter_segments_from_outer_loop(
        loops3[outer_index],
        p0_on_face,
        direction,
        interior_normal,
        depth_from_interior,
        base_z,
        start_d,
    )

    host = Host3()
    host.wall = wall
    host.wall_id = wall.Id
    host.level = level
    host.base_z = base_z
    host.start = target_origin + direction * start_d
    host.direction = direction
    host.interior_normal = interior_normal
    host.length = length
    host.outline_2d = outer_local
    host.openings = openings
    host.perimeter_segments = perimeter_segments
    host.audit = {
        "wall": element_id_text(wall.Id),
        "location_len": curve.Length,
        "face_len": length,
        "start_shift": start_d,
        "depth_from_interior": depth_from_interior,
        "layer": layer_label,
        "loops": len(loops2),
        "face_openings": len(face_openings),
        "hosted_openings": len(hosted_openings),
        "final_openings": len(openings),
        "perimeter_segments": len(perimeter_segments),
    }
    return host


def largest_interior_face(wall):
    try:
        refs = DB.HostObjectUtils.GetSideFaces(wall, DB.ShellLayerType.Interior)
    except Exception:
        refs = []
    best = None
    best_area = 0.0
    for ref in refs:
        try:
            face = wall.GetGeometryObjectFromReference(ref)
        except Exception:
            face = None
        if face is None:
            continue
        area = float(getattr(face, "Area", 0.0) or 0.0)
        if best is None or area > best_area:
            best = face
            best_area = area
    return best


def face_curve_loops(face):
    try:
        curve_loops = face.GetEdgesAsCurveLoops()
    except Exception:
        curve_loops = None
    if curve_loops is None:
        return []
    loops = []
    for curve_loop in curve_loops:
        points = []
        for curve in curve_loop:
            try:
                tess = curve.Tessellate()
            except Exception:
                tess = []
            for point in tess:
                if points and same_xyz(points[-1], point):
                    continue
                points.append(point)
        if len(points) >= 3:
            loops.append(points)
    return loops


def interior_face_normal(face, wall):
    normal = None
    try:
        normal = face.ComputeNormal(DB.UV(0.5, 0.5))
    except Exception:
        try:
            normal = face.FaceNormal
        except Exception:
            normal = None
    normal = horizontal_unit(normal)
    if normal is None:
        return None
    exterior = horizontal_unit(getattr(wall, "Orientation", None))
    if exterior is not None:
        try:
            if normal.DotProduct(exterior) > 0.0:
                normal = DB.XYZ(-normal.X, -normal.Y, 0.0)
        except Exception:
            pass
    return normal


def target_depth_from_interior(doc, wall):
    compound = None
    try:
        type_elem = doc.GetElement(wall.GetTypeId())
        compound = type_elem.GetCompoundStructure()
    except Exception:
        compound = None
    if compound is None:
        return 0.0, "no compound structure"
    try:
        layers = list(compound.GetLayers())
    except Exception:
        layers = []
    if not layers:
        return 0.0, "no compound layers"

    widths = []
    total = 0.0
    for layer in layers:
        width = max(0.0, float(getattr(layer, "Width", 0.0) or 0.0))
        widths.append(width)
        total += width
    if total <= 1e-9:
        return 0.0, "zero width compound"

    start_depths = []
    cursor = 0.0
    for width in widths:
        start_depths.append(cursor)
        cursor += width

    try:
        structural_index = int(compound.StructuralMaterialIndex)
    except Exception:
        structural_index = -1
    if 0 <= structural_index < len(widths) and widths[structural_index] > 1e-9:
        center = start_depths[structural_index] + widths[structural_index] * 0.5
        return max(0.0, total - center), "structural layer {0}".format(structural_index)

    try:
        first_core = int(compound.GetFirstCoreLayerIndex())
        last_core = int(compound.GetLastCoreLayerIndex())
    except Exception:
        first_core = -1
        last_core = -1
    if 0 <= first_core <= last_core < len(widths):
        core_start = start_depths[first_core]
        core_end = start_depths[last_core] + widths[last_core]
        center = (core_start + core_end) * 0.5
        return max(0.0, total - center), "core center"

    thickest = 0
    for index in range(len(widths)):
        if widths[index] > widths[thickest]:
            thickest = index
    center = start_depths[thickest] + widths[thickest] * 0.5
    return max(0.0, total - center), "thickest layer {0}".format(thickest)


def loops_local_2d(loops3, origin, direction):
    result = []
    for loop in loops3:
        local = []
        for pt in loop:
            d = (pt - origin).DotProduct(direction)
            candidate = (d, pt.Z)
            if local:
                prev = local[-1]
                if abs(prev[0] - candidate[0]) < 1e-7 and abs(prev[1] - candidate[1]) < 1e-7:
                    continue
            local.append(candidate)
        if len(local) >= 3:
            result.append(local)
    return result


def perimeter_segments_from_outer_loop(points, origin, direction, interior_normal,
                                       depth_from_interior, base_z, start_d):
    if len(points) < 2:
        return []

    z_values = [point.Z for point in points]
    low_z = min(z_values)
    segments = []
    count = len(points)
    for index in range(count):
        p0_face = points[index]
        p1_face = points[(index + 1) % count]
        if same_xyz(p0_face, p1_face):
            continue
        p0 = offset_from_interior_face(p0_face, interior_normal, depth_from_interior)
        p1 = offset_from_interior_face(p1_face, interior_normal, depth_from_interior)
        if p0.DistanceTo(p1) < MIN_LEN:
            continue

        horizontal = math.sqrt(
            (p1.X - p0.X) * (p1.X - p0.X) +
            (p1.Y - p0.Y) * (p1.Y - p0.Y)
        )
        vertical = abs(p1.Z - p0.Z)
        if horizontal < 0.5 / 12.0 and vertical > 6.0 / 12.0:
            kind = "side"
        elif max(p0.Z, p1.Z) <= low_z + 3.0 / 12.0:
            kind = "bottom"
        else:
            kind = "top"

        d0 = (p0_face - origin).DotProduct(direction) - start_d
        d1 = (p1_face - origin).DotProduct(direction) - start_d
        segments.append(FaceSegment3(kind, p0, p1, d0, d1))
    return segments


def offset_from_interior_face(point, interior_normal, depth_from_interior):
    target = point - interior_normal * depth_from_interior
    return DB.XYZ(target.X, target.Y, target.Z)


def largest_loop_index(loops):
    best_index = None
    best_area = 0.0
    for index, loop in enumerate(loops):
        area = abs(poly_area(loop))
        if best_index is None or area > best_area:
            best_index = index
            best_area = area
    return best_index


def opening_from_loop(loop, start_d, length, base_z):
    left = min(d for d, _ in loop) - start_d
    right = max(d for d, _ in loop) - start_d
    sill_abs = min(z for _, z in loop)
    head_abs = max(z for _, z in loop)
    left = max(0.0, left)
    right = min(length, right)
    if left <= STUD_T * 2.0 or right >= length - STUD_T * 2.0:
        return None
    if right - left < 6.0 / 12.0:
        return None
    if head_abs - sill_abs < 6.0 / 12.0:
        return None
    return Opening3(left, right, sill_abs - base_z, head_abs - base_z, sill_abs > base_z + PLATE_T * 2.0)


def hosted_openings_3(doc, wall, origin, direction, start_d, end_d, base_z):
    try:
        ids = wall.FindInserts(True, False, False, False)
    except Exception:
        ids = []
    openings = []
    for element_id in ids:
        elem = doc.GetElement(element_id)
        opening = opening_from_insert_3(elem, origin, direction, start_d, end_d, base_z)
        if opening is not None:
            openings.append(opening)
    return openings


def opening_from_insert_3(elem, origin, direction, start_d, end_d, base_z):
    if elem is None:
        return None
    try:
        if isinstance(elem, DB.Opening):
            rect = elem.BoundaryRect
            if rect and len(rect) >= 2:
                d0 = (rect[0] - origin).DotProduct(direction)
                d1 = (rect[1] - origin).DotProduct(direction)
                return opening_abs(min(d0, d1), max(d0, d1), min(rect[0].Z, rect[1].Z), max(rect[0].Z, rect[1].Z), start_d, end_d, base_z)
    except Exception:
        pass

    is_window = category_matches(elem, DB.BuiltInCategory.OST_Windows)
    is_door = category_matches(elem, DB.BuiltInCategory.OST_Doors)
    if not is_window and not is_door:
        return None

    point = getattr(getattr(elem, "Location", None), "Point", None)
    if point is None:
        return None
    center = (point - origin).DotProduct(direction)
    width = read_opening_dim(elem, "width")
    height = read_opening_dim(elem, "height")
    sill = read_sill(elem) if is_window else 0.0
    return opening_abs(center - width * 0.5, center + width * 0.5, base_z + sill, base_z + sill + height, start_d, end_d, base_z)


def opening_abs(left_abs, right_abs, sill_abs, head_abs, start_d, end_d, base_z):
    length = end_d - start_d
    left = max(0.0, left_abs - start_d)
    right = min(length, right_abs - start_d)
    if right - left < 6.0 / 12.0:
        return None
    if head_abs - sill_abs < 6.0 / 12.0:
        return None
    return Opening3(left, right, sill_abs - base_z, head_abs - base_z, sill_abs > base_z + PLATE_T * 2.0)


def read_opening_dim(elem, kind):
    if kind == "width":
        builtins = ("FAMILY_ROUGH_WIDTH_PARAM", "DOOR_ROUGH_WIDTH", "WINDOW_ROUGH_WIDTH", "FAMILY_WIDTH_PARAM", "DOOR_WIDTH", "WINDOW_WIDTH", "GENERIC_WIDTH")
        names = ("Rough Width", "Rough Opening Width", "Width")
        default = 3.0
    else:
        builtins = ("FAMILY_ROUGH_HEIGHT_PARAM", "DOOR_ROUGH_HEIGHT", "WINDOW_ROUGH_HEIGHT", "FAMILY_HEIGHT_PARAM", "DOOR_HEIGHT", "WINDOW_HEIGHT", "GENERIC_HEIGHT")
        names = ("Rough Height", "Rough Opening Height", "Height")
        default = 6.667
    for source in (elem, getattr(elem, "Symbol", None)):
        if source is None:
            continue
        for builtin_name in builtins:
            param_id = getattr(DB.BuiltInParameter, builtin_name, None)
            if param_id is None:
                continue
            value = param_double(source, param_id, None)
            if value is not None and value > 0.0:
                return value
        for name in names:
            value = lookup_double(source, name, None)
            if value is not None and value > 0.0:
                return value
    return default


def read_sill(elem):
    builtins = ("INSTANCE_SILL_HEIGHT_PARAM", "FAMILY_SILL_HEIGHT_PARAM")
    for source in (elem, getattr(elem, "Symbol", None)):
        if source is None:
            continue
        for builtin_name in builtins:
            param_id = getattr(DB.BuiltInParameter, builtin_name, None)
            if param_id is None:
                continue
            value = param_double(source, param_id, None)
            if value is not None and value >= 0.0:
                return value
        value = lookup_double(source, "Sill Height", None)
        if value is not None and value >= 0.0:
            return value
    return 3.0


def merge_openings(hosted, face, length):
    merged = []
    for source in (hosted, face):
        for op in source:
            left = max(0.0, op.left)
            right = min(length, op.right)
            if right - left < 6.0 / 12.0:
                continue
            duplicate = False
            for existing in merged:
                if abs(existing.left - left) < STUD_T and abs(existing.right - right) < STUD_T:
                    duplicate = True
                    break
            if not duplicate:
                merged.append(Opening3(left, right, op.sill, op.head, op.is_window))
    merged.sort(key=lambda op: op.left)
    return merged


def build_members_3(host, options):
    members = []
    occupied = set()
    members.extend(perimeter_members(host, options, occupied))
    members.extend(opening_members(host, options, occupied))
    members.extend(regular_studs(host, options, occupied))
    return [member for member in members if member is not None]


def perimeter_members(host, options, occupied):
    result = []
    for segment in host.perimeter_segments:
        if segment.kind == "bottom":
            result.extend(perimeter_bottom_plates(host, options, segment))
        elif segment.kind == "top":
            result.extend(perimeter_top_plates(options, segment))
        elif segment.kind == "side":
            member, d = perimeter_side_stud(host, options, segment)
            if member is not None:
                result.append(member)
                occupied.add(round(d, 4))
    return result


def perimeter_bottom_plates(host, options, segment):
    result = []
    gaps = [(op.left, op.right) for op in host.openings if not op.is_window]
    d_min = min(segment.d0, segment.d1)
    d_max = max(segment.d0, segment.d1)
    for plate_index in range(options["bottom_count"]):
        z_delta = PLATE_T * (plate_index + 0.5)
        for left, right in split_segments(d_min, d_max, gaps):
            p0 = point_at_segment_d(segment, left, z_delta)
            p1 = point_at_segment_d(segment, right, z_delta)
            result.append(member_from_points("bottom_plate", p0, p1, options["plate"], False, -math.pi / 2.0))
    return result


def perimeter_top_plates(options, segment):
    result = []
    for plate_index in range(options["top_count"]):
        z_delta = -(options["top_count"] - plate_index - 0.5) * PLATE_T
        p0 = add_z(segment.p0, z_delta)
        p1 = add_z(segment.p1, z_delta)
        result.append(member_from_points("top_plate", p0, p1, options["plate"], False, -math.pi / 2.0))
    return result


def perimeter_side_stud(host, options, segment):
    low = segment.p0
    high = segment.p1
    if low.Z > high.Z:
        low, high = high, low
    bottom_z = low.Z + options["bottom_count"] * PLATE_T
    top_z = high.Z - options["top_count"] * PLATE_T
    if top_z - bottom_z < MIN_LEN:
        return None, (segment.d0 + segment.d1) * 0.5
    p0 = point_at_segment_z(segment, bottom_z)
    p1 = point_at_segment_z(segment, top_z)
    d = (segment.d0 + segment.d1) * 0.5
    return member_from_points("side_stud", p0, p1, options["stud"], True, segment_angle_at_d(host, d)), d


def opening_members(host, options, occupied):
    result = []
    header_depth = symbol_depth(options["header"], DEFAULT_HEADER_DEPTH)
    for op in host.openings:
        left = op.left
        right = op.right
        for d in (left - STUD_T * 1.5, right + STUD_T * 1.5):
            add_vertical_stud_at_d(result, occupied, "king_stud", host, options, d)
        for d in (left - STUD_T * 0.5, right + STUD_T * 0.5):
            if d <= 0.0 or d >= host.length or near(d, occupied, STUD_T):
                continue
            member = vertical_member_at_d(
                "jack_stud",
                host,
                options,
                d,
                None,
                host.base_z + op.head,
                options["stud"],
            )
            if member is not None:
                result.append(member)
                occupied.add(round(d, 4))

        h_center = op.head + header_depth * 0.5
        span_left = max(0.0, left - STUD_T)
        span_right = min(host.length, right + STUD_T)
        for header_index in range(2):
            lateral = (header_index - 0.5) * STUD_T
            result.append(
                beam_at_d(
                    "header",
                    host,
                    span_left,
                    span_right,
                    host.base_z + h_center,
                    options["header"],
                    0.0,
                    lateral,
                )
            )

        if op.is_window and op.sill > stud_bottom(options):
            sill_h = op.sill - PLATE_T * 0.5
            result.append(
                beam_at_d(
                    "sill_plate",
                    host,
                    left,
                    right,
                    host.base_z + sill_h,
                    options["plate"],
                    -math.pi / 2.0,
                    0.0,
                )
            )

        header_top = op.head + header_depth
        top_abs = top_z_at_d(host, (left + right) * 0.5)
        if top_abs is not None:
            top = top_abs - host.base_z - options["top_count"] * PLATE_T
            if header_top < top:
                result.extend(cripple_studs(host, options, left, right, header_top, top, occupied))
        if op.is_window:
            below_top = op.sill - PLATE_T
            if below_top > stud_bottom(options):
                result.extend(cripple_studs(host, options, left, right, stud_bottom(options), below_top, occupied))
    return result


def regular_studs(host, options, occupied):
    result = []
    spacing = options["spacing"]
    for segment in bottom_segments(host):
        d_min = max(0.0, min(segment.d0, segment.d1))
        d_max = min(host.length, max(segment.d0, segment.d1))
        if d_max - d_min < MIN_LEN:
            continue
        d = next_spacing_station(d_min, spacing)
        while d < d_max - STUD_T * 0.5:
            if not in_opening_zone(host, d) and not near(d, occupied, STUD_T):
                member = vertical_member_on_bottom_segment(
                    "stud",
                    host,
                    options,
                    segment,
                    d,
                    None,
                    None,
                    options["stud"],
                )
                if member is not None:
                    result.append(member)
                    occupied.add(round(d, 4))
            d += spacing
    return result


def add_vertical_stud_at_d(result, occupied, kind, host, options, d):
    if d <= 0.0 or d >= host.length or near(d, occupied, STUD_T):
        return
    member = vertical_member_at_d(
        kind,
        host,
        options,
        d,
        None,
        None,
        options["stud"],
    )
    if member is not None:
        result.append(member)
        occupied.add(round(d, 4))


def vertical_member_at_d(kind, host, options, d, bottom_abs_z, top_abs_z, symbol):
    segment = segment_at_d(bottom_segments(host), d)
    if segment is None:
        return None
    return vertical_member_on_bottom_segment(
        kind,
        host,
        options,
        segment,
        d,
        bottom_abs_z,
        top_abs_z,
        symbol,
    )


def vertical_member_on_bottom_segment(kind, host, options, segment, d,
                                      bottom_abs_z, top_abs_z, symbol):
    base_point = point_at_segment_d(segment, d, 0.0)
    if base_point is None:
        return None
    wall_bottom = base_point.Z
    wall_top = top_z_from_perimeter(host, d)
    if wall_top is None:
        bounds = wall_z_bounds_at_d(host, d)
        if bounds is None:
            return None
        wall_bottom, wall_top = bounds
    framed_bottom = wall_bottom + options["bottom_count"] * PLATE_T
    framed_top = wall_top - options["top_count"] * PLATE_T
    if bottom_abs_z is None or bottom_abs_z < framed_bottom:
        bottom_abs_z = framed_bottom
    if top_abs_z is None or top_abs_z > framed_top:
        top_abs_z = framed_top
    if top_abs_z - bottom_abs_z < MIN_LEN:
        return None
    start = DB.XYZ(base_point.X, base_point.Y, bottom_abs_z)
    end = DB.XYZ(base_point.X, base_point.Y, top_abs_z)
    angle = segment_angle(segment)
    return member_from_points(kind, start, end, symbol, True, angle)


def beam_at_d(kind, host, d0, d1, z_abs, symbol, rotation, lateral=0.0):
    p0 = bottom_point_at_d(host, d0)
    p1 = bottom_point_at_d(host, d1)
    if p0 is None or p1 is None:
        return None
    start = DB.XYZ(p0.X, p0.Y, z_abs)
    end = DB.XYZ(p1.X, p1.Y, z_abs)
    if abs(lateral) > 1e-9:
        start = start + host.interior_normal * lateral
        end = end + host.interior_normal * lateral
    return member_from_points(kind, start, end, symbol, False, rotation)


def bottom_segments(host):
    return [segment for segment in host.perimeter_segments if segment.kind == "bottom"]


def next_spacing_station(start_d, spacing):
    if spacing <= 0.0:
        return start_d
    return math.ceil((start_d + STUD_T * 0.5) / spacing) * spacing


def bottom_point_at_d(host, d):
    segment = segment_at_d(bottom_segments(host), d)
    if segment is None:
        return None
    return point_at_segment_d(segment, d, 0.0)


def top_z_at_d(host, d):
    top_z = top_z_from_perimeter(host, d)
    if top_z is not None:
        return top_z
    bounds = wall_z_bounds_at_d(host, d)
    if bounds is None:
        return None
    return bounds[1]


def top_z_from_perimeter(host, d):
    candidates = []
    for segment in host.perimeter_segments:
        if segment.kind != "top":
            continue
        if not segment_contains_d(segment, d, STUD_T):
            continue
        point = point_at_segment_d(segment, d, 0.0)
        if point is not None:
            candidates.append(point.Z)
    if candidates:
        return max(candidates)
    return None


def wall_z_bounds_at_d(host, d):
    intersections = []
    outline = getattr(host, "outline_2d", None) or []
    if len(outline) < 3:
        return None
    tol = 1e-7
    count = len(outline)
    for index in range(count):
        d0, z0 = outline[index]
        d1, z1 = outline[(index + 1) % count]

        if abs(d1 - d0) < tol:
            if abs(d - d0) <= STUD_T:
                intersections.append(z0)
                intersections.append(z1)
            continue

        min_d = min(d0, d1) - tol
        max_d = max(d0, d1) + tol
        if d < min_d or d > max_d:
            continue
        t = (d - d0) / (d1 - d0)
        if t < -tol or t > 1.0 + tol:
            continue
        t = max(0.0, min(1.0, t))
        intersections.append(z0 + (z1 - z0) * t)

    values = unique_sorted(intersections)
    if len(values) < 2:
        return None
    return values[0], values[-1]


def unique_sorted(values):
    result = []
    for value in sorted(values):
        if not result or abs(value - result[-1]) > 1e-6:
            result.append(value)
    return result


def segment_angle_at_d(host, d):
    segment = segment_at_d(bottom_segments(host), d)
    if segment is None:
        return wall_angle(host.direction)
    return segment_angle(segment)


def segment_at_d(segments, d):
    best = None
    best_gap = None
    for segment in segments:
        if segment_contains_d(segment, d, STUD_T):
            return segment
        d0 = min(segment.d0, segment.d1)
        d1 = max(segment.d0, segment.d1)
        gap = min(abs(d - d0), abs(d - d1))
        if best is None or gap < best_gap:
            best = segment
            best_gap = gap
    return best


def segment_contains_d(segment, d, tolerance):
    d0 = min(segment.d0, segment.d1)
    d1 = max(segment.d0, segment.d1)
    return d0 - tolerance <= d <= d1 + tolerance


def segment_angle(segment):
    return math.atan2(segment.p1.Y - segment.p0.Y, segment.p1.X - segment.p0.X)


def cripple_studs(host, options, left, right, bottom, top, occupied):
    result = []
    if top - bottom < MIN_LEN:
        return result
    d = options["spacing"]
    while d < right - STUD_T * 0.5:
        if left + STUD_T * 0.5 < d < right - STUD_T * 0.5:
            member = vertical_member_at_d(
                "cripple_stud",
                host,
                options,
                d,
                host.base_z + bottom,
                host.base_z + top,
                options["stud"],
            )
            if member is not None:
                result.append(member)
        d += options["spacing"]
    return result


def member_from_points(kind, start, end, symbol, is_column, rotation):
    if start is None or end is None:
        return None
    if start.DistanceTo(end) < MIN_LEN:
        return None
    return Member3(kind, start, end, symbol, is_column, rotation)


def point_at_segment_d(segment, d, z_delta):
    denom = segment.d1 - segment.d0
    if abs(denom) < 1e-9:
        t = 0.0
    else:
        t = (d - segment.d0) / denom
    t = max(0.0, min(1.0, t))
    return DB.XYZ(
        segment.p0.X + (segment.p1.X - segment.p0.X) * t,
        segment.p0.Y + (segment.p1.Y - segment.p0.Y) * t,
        segment.p0.Z + (segment.p1.Z - segment.p0.Z) * t + z_delta,
    )


def point_at_segment_z(segment, z):
    denom = segment.p1.Z - segment.p0.Z
    if abs(denom) < 1e-9:
        t = 0.0
    else:
        t = (z - segment.p0.Z) / denom
    t = max(0.0, min(1.0, t))
    return DB.XYZ(
        segment.p0.X + (segment.p1.X - segment.p0.X) * t,
        segment.p0.Y + (segment.p1.Y - segment.p0.Y) * t,
        z,
    )


def add_z(point, delta):
    return DB.XYZ(point.X, point.Y, point.Z + delta)


def stud_bottom(options):
    return options["bottom_count"] * PLATE_T


def place_members_3(doc, host, members):
    placed = 0
    for member in members:
        instance = None
        if member.is_column:
            instance = place_column_3(doc, host.level, member)
        else:
            instance = place_beam_3(doc, host.level, member)
        if instance is None:
            continue
        tag_wf3(instance, host, member)
        placed += 1
    return placed


def place_column_3(doc, level, member):
    try:
        instance = doc.Create.NewFamilyInstance(member.start, member.symbol, level, StructuralType.Column)
    except Exception:
        return None
    set_element_id(instance, DB.BuiltInParameter.FAMILY_BASE_LEVEL_PARAM, level.Id)
    set_element_id(instance, DB.BuiltInParameter.FAMILY_TOP_LEVEL_PARAM, level.Id)
    set_double(instance, DB.BuiltInParameter.FAMILY_BASE_LEVEL_OFFSET_PARAM, member.start.Z - level.Elevation)
    set_double(instance, DB.BuiltInParameter.FAMILY_TOP_LEVEL_OFFSET_PARAM, member.end.Z - level.Elevation)
    rotate_element_z(doc, instance, member.start, member.rotation)
    return instance


def place_beam_3(doc, level, member):
    try:
        line = DB.Line.CreateBound(member.start, member.end)
        instance = doc.Create.NewFamilyInstance(line, member.symbol, level, StructuralType.Beam)
    except Exception:
        return None
    center_framing(instance)
    set_double(instance, DB.BuiltInParameter.STRUCTURAL_BEND_DIR_ANGLE, member.rotation)
    try:
        from Autodesk.Revit.DB.Structure import StructuralFramingUtils
        StructuralFramingUtils.DisallowJoinAtEnd(instance, 0)
        StructuralFramingUtils.DisallowJoinAtEnd(instance, 1)
    except Exception:
        pass
    return instance


def tag_wf3(instance, host, member):
    try:
        param = instance.get_Parameter(DB.BuiltInParameter.ALL_MODEL_INSTANCE_COMMENTS)
        if param is None or param.IsReadOnly:
            return
        value = "{0}wall={1}|member={2}".format(TRACK_PREFIX, element_id_text(host.wall_id), member.kind)
        param.Set(value)
    except Exception:
        pass


def build_report(wall_total, skipped, placed_total, rows):
    text = (
        "## Wall Framing 3.0 Complete\n"
        "- **Engine:** {0}\n"
        "- **Standalone file:** Yes\n"
        "- **Existing WF project libraries imported:** No\n"
        "- **Existing members deleted:** 0\n"
        "- **Walls selected:** {1}\n"
        "- **Walls skipped:** {2}\n"
        "- **Members placed:** {3}\n\n"
        "### Source Geometry Audit\n"
        "| Wall Id | Location Len | Face Len | Start Shift | Interior Depth | Layer | Loops | Perimeter Segs | Face Openings | Hosted Openings | Final Openings | Placed |\n"
        "| --- | ---: | ---: | ---: | ---: | --- | ---: | ---: | ---: | ---: | ---: | ---: |\n"
    ).format(ENGINE_NAME, wall_total, skipped, placed_total)
    for audit, placed in rows:
        text += "| {0} | {1:.3f} | {2:.3f} | {3:.3f} | {4:.3f} | {5} | {6} | {7} | {8} | {9} | {10} | {11} |\n".format(
            audit.get("wall", "?"),
            float(audit.get("location_len") or 0.0),
            float(audit.get("face_len") or 0.0),
            float(audit.get("start_shift") or 0.0),
            float(audit.get("depth_from_interior") or 0.0),
            audit.get("layer", "?"),
            audit.get("loops", 0),
            audit.get("perimeter_segments", 0),
            audit.get("face_openings", 0),
            audit.get("hosted_openings", 0),
            audit.get("final_openings", 0),
            placed,
        )
    return text


def activate_symbol(doc, symbol):
    try:
        if not symbol.IsActive:
            symbol.Activate()
    except Exception:
        pass


def symbol_label(symbol):
    family = first_text(
        symbol,
        (
            "SYMBOL_FAMILY_NAME_PARAM",
            "ALL_MODEL_FAMILY_NAME",
        ),
    )
    type_name = first_text(
        symbol,
        (
            "SYMBOL_NAME_PARAM",
            "ALL_MODEL_TYPE_NAME",
        ),
    )

    if not family:
        family = safe_attr_text(symbol, "FamilyName")
    if not family:
        family_obj = safe_attr(symbol, "Family")
        family = element_name(family_obj)
    if not type_name:
        type_name = element_name(symbol)

    if family and type_name:
        return family + " : " + type_name
    if type_name:
        return type_name
    if family:
        return family
    return "Element " + element_id_text(symbol.Id)


def first_text(element, builtin_names):
    for builtin_name in builtin_names:
        param_id = getattr(DB.BuiltInParameter, builtin_name, None)
        if param_id is None:
            continue
        value = param_text(element, param_id)
        if value:
            return value
    return None


def param_text(element, param_id):
    try:
        param = element.get_Parameter(param_id)
    except Exception:
        param = None
    if param is None or not param.HasValue:
        return None
    try:
        value = param.AsString()
        if value:
            return str(value)
    except Exception:
        pass
    try:
        value = param.AsValueString()
        if value:
            return str(value)
    except Exception:
        pass
    return None


def safe_attr(element, name):
    if element is None:
        return None
    try:
        return getattr(element, name)
    except Exception:
        return None


def safe_attr_text(element, name):
    value = safe_attr(element, name)
    if value is None:
        return None
    try:
        text = str(value)
    except Exception:
        return None
    if text:
        return text
    return None


def element_name(element):
    if element is None:
        return None
    try:
        value = DB.Element.Name.GetValue(element)
        if value:
            return str(value)
    except Exception:
        pass
    return safe_attr_text(element, "Name")


def center_framing(instance):
    set_int(instance, DB.BuiltInParameter.YZ_JUSTIFICATION, 0)
    set_int(instance, DB.BuiltInParameter.Y_JUSTIFICATION, 2)
    set_int(instance, DB.BuiltInParameter.Z_JUSTIFICATION, 2)
    for param_id in (
        DB.BuiltInParameter.Y_OFFSET_VALUE,
        DB.BuiltInParameter.Z_OFFSET_VALUE,
        DB.BuiltInParameter.START_Y_OFFSET_VALUE,
        DB.BuiltInParameter.END_Y_OFFSET_VALUE,
        DB.BuiltInParameter.START_Z_OFFSET_VALUE,
        DB.BuiltInParameter.END_Z_OFFSET_VALUE,
    ):
        set_double(instance, param_id, 0.0)


def set_int(element, param_id, value):
    try:
        param = element.get_Parameter(param_id)
        if param is not None and not param.IsReadOnly:
            param.Set(int(value))
    except Exception:
        pass


def set_double(element, param_id, value):
    try:
        param = element.get_Parameter(param_id)
        if param is not None and not param.IsReadOnly:
            param.Set(float(value))
    except Exception:
        pass


def set_element_id(element, param_id, value):
    try:
        param = element.get_Parameter(param_id)
        if param is not None and not param.IsReadOnly:
            param.Set(value)
    except Exception:
        pass


def rotate_element_z(doc, element, base, angle):
    if abs(angle) < 1e-9:
        return
    try:
        axis = DB.Line.CreateBound(base, base + DB.XYZ.BasisZ)
        DB.ElementTransformUtils.RotateElement(doc, element.Id, axis, angle)
    except Exception:
        pass


def symbol_depth(symbol, default):
    for name in ("d", "Depth", "Height"):
        value = lookup_double(symbol, name, None)
        if value is not None and value > 0.0:
            return value
    return default


def param_double(element, param_id, default):
    try:
        param = element.get_Parameter(param_id)
        if param is not None and param.HasValue:
            return param.AsDouble()
    except Exception:
        pass
    return default


def lookup_double(element, name, default):
    try:
        param = element.LookupParameter(name)
        if param is not None and param.HasValue:
            return param.AsDouble()
    except Exception:
        pass
    return default


def category_matches(element, target):
    category = getattr(element, "Category", None)
    if category is None:
        return False
    current = getattr(category.Id, "IntegerValue", getattr(category.Id, "Value", None))
    return current == int(target)


def first_point(loops):
    for loop in loops:
        if loop:
            return loop[0]
    return None


def same_xyz(a, b):
    return abs(a.X - b.X) < 1e-7 and abs(a.Y - b.Y) < 1e-7 and abs(a.Z - b.Z) < 1e-7


def normalize_xyz(vector):
    try:
        length = math.sqrt(vector.X * vector.X + vector.Y * vector.Y + vector.Z * vector.Z)
    except Exception:
        return None
    if length < 1e-9:
        return None
    return DB.XYZ(vector.X / length, vector.Y / length, vector.Z / length)


def horizontal_unit(vector):
    if vector is None:
        return None
    try:
        length = math.sqrt(vector.X * vector.X + vector.Y * vector.Y)
    except Exception:
        return None
    if length < 1e-9:
        return None
    return DB.XYZ(vector.X / length, vector.Y / length, 0.0)


def wall_angle(direction):
    return math.atan2(direction.Y, direction.X)


def poly_area(points):
    area = 0.0
    for i in range(len(points)):
        x0, y0 = points[i]
        x1, y1 = points[(i + 1) % len(points)]
        area += x0 * y1 - x1 * y0
    return area * 0.5


def split_segments(start, end, gaps):
    segments = [(start, end)]
    for gap_start, gap_end in gaps:
        next_segments = []
        for left, right in segments:
            if gap_end <= left or gap_start >= right:
                next_segments.append((left, right))
                continue
            if gap_start > left:
                next_segments.append((left, gap_start))
            if gap_end < right:
                next_segments.append((gap_end, right))
        segments = next_segments
    return [(left, right) for left, right in segments if right - left >= MIN_LEN]


def near(value, occupied, tolerance):
    for existing in occupied:
        if abs(existing - value) < tolerance:
            return True
    return False


def in_opening_zone(host, d):
    for op in host.openings:
        if op.left - STUD_T * 3.0 <= d <= op.right + STUD_T * 3.0:
            return True
    return False


def element_id_text(element_id):
    if element_id is None:
        return "?"
    value = getattr(element_id, "IntegerValue", getattr(element_id, "Value", None))
    if value is None:
        return str(element_id)
    return str(value)


if __name__ == "__main__":
    main()
