# -*- coding: utf-8 -*-
"""Ownership tracking for generated framing members."""

TRACKING_PREFIX = "WF_FRAME|"


def tag_instance(instance, host_info, member):
    """Tag a generated instance so update/delete can find it safely."""
    from Autodesk.Revit.DB import BuiltInParameter

    param = instance.get_Parameter(BuiltInParameter.ALL_MODEL_INSTANCE_COMMENTS)
    if param is None or param.IsReadOnly:
        return False

    tracking_line = _build_tracking_line(host_info, member)
    existing = param.AsString() or ""

    lines = []
    for line in existing.splitlines():
        if not line.startswith(TRACKING_PREFIX):
            lines.append(line)
    lines.append(tracking_line)

    try:
        param.Set("\n".join(lines))
        return True
    except Exception:
        return False


def get_tracking_data(element):
    """Read tracking metadata from a generated framing element."""
    from Autodesk.Revit.DB import BuiltInParameter

    param = element.get_Parameter(BuiltInParameter.ALL_MODEL_INSTANCE_COMMENTS)
    if param is None:
        return None

    text = param.AsString() or ""
    for line in text.splitlines():
        if line.startswith(TRACKING_PREFIX):
            return _parse_tracking_line(line)
    return None


def get_tracked_members_for_hosts(doc, hosts):
    """Return tracked structural framing instances for the given hosts."""
    from Autodesk.Revit.DB import BuiltInCategory, FilteredElementCollector

    host_keys = set()
    for host in hosts:
        host_key = host_key_for_element(host)
        if host_key is not None:
            host_keys.add(host_key)

    if not host_keys:
        return []

    collector = (
        FilteredElementCollector(doc)
        .OfCategory(BuiltInCategory.OST_StructuralFraming)
        .WhereElementIsNotElementType()
    )

    matches = []
    for element in collector:
        tracking = get_tracking_data(element)
        if tracking is None:
            continue
        key = tracking.get("kind"), tracking.get("host")
        if key in host_keys:
            matches.append(element)
    return matches


def host_key_for_element(element):
    """Return the tracking key for a selected host element."""
    from Autodesk.Revit.DB import BuiltInCategory, Floor, RoofBase, Wall

    if isinstance(element, Wall):
        return ("wall", str(_element_id_value(element.Id)))
    if isinstance(element, Floor):
        return ("floor", str(_element_id_value(element.Id)))
    if _category_matches(element, BuiltInCategory.OST_Ceilings):
        return ("ceiling", str(_element_id_value(element.Id)))
    if isinstance(element, RoofBase):
        return ("roof", str(_element_id_value(element.Id)))
    return None


def get_nearby_structural_framing(doc, element, tolerance=0.5):
    """Fallback collector for legacy untracked framing near a host."""
    from Autodesk.Revit.DB import (
        BoundingBoxIntersectsFilter,
        BuiltInCategory,
        FilteredElementCollector,
        Outline,
        XYZ,
    )

    bounding_box = element.get_BoundingBox(None)
    if bounding_box is None:
        return []

    outline = Outline(
        XYZ(
            bounding_box.Min.X - tolerance,
            bounding_box.Min.Y - tolerance,
            bounding_box.Min.Z - tolerance,
        ),
        XYZ(
            bounding_box.Max.X + tolerance,
            bounding_box.Max.Y + tolerance,
            bounding_box.Max.Z + tolerance,
        ),
    )
    collector = (
        FilteredElementCollector(doc)
        .OfCategory(BuiltInCategory.OST_StructuralFraming)
        .WhereElementIsNotElementType()
        .WherePasses(BoundingBoxIntersectsFilter(outline))
    )
    return list(collector)


def _build_tracking_line(host_info, member):
    """Create a single-line tracking record."""
    values = [
        ("kind", getattr(member, "host_kind", None) or getattr(host_info, "kind", None)),
        ("host", _element_id_value(getattr(member, "host_id", None) or getattr(host_info, "element_id", None))),
        ("member", member.member_type),
        ("layer", getattr(member, "layer_index", None)),
    ]

    parts = [TRACKING_PREFIX.rstrip("|")]
    for key, value in values:
        if value is None or value == "":
            continue
        parts.append("%s=%s" % (key, value))
    return "|".join(parts)


def _parse_tracking_line(line):
    """Parse a tracking line into a dictionary."""
    result = {}
    parts = line.split("|")
    for part in parts[1:]:
        if "=" not in part:
            continue
        key, value = part.split("=", 1)
        result[key] = value
    return result


def _category_matches(element, category_id):
    """Return True when the element belongs to the given category."""
    category = getattr(element, "Category", None)
    if category is None:
        return False

    current_id = getattr(category.Id, "IntegerValue", getattr(category.Id, "Value", None))
    target_id = int(category_id)
    return current_id == target_id


def _element_id_value(element_id):
    """Return a numeric id for an ElementId or raw integer."""
    if element_id is None:
        return None
    if isinstance(element_id, (int, long)):
        return element_id
    return getattr(element_id, "IntegerValue", getattr(element_id, "Value", None))