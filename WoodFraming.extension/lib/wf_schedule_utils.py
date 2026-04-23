# -*- coding: utf-8 -*-
"""Shared parameter, schedule, and sheathing helpers for Wood Framing."""

import math
import os
from collections import OrderedDict

from pyrevit import DB, revit

from wf_config import FramingConfig
from wf_host import (
    analyze_ceiling_host,
    analyze_floor_host,
    analyze_roof_host,
    analyze_wall_host,
)
from wf_tracking import get_tracking_data


BOM_SCHEDULE_NAME = "WF - BOM by Host"
SHEATHING_SCHEDULE_NAME = "WF - Sheathing by Host"
SHARED_PARAM_GROUP_NAME = "WoodFraming"
SHARED_PARAM_FILE = os.path.join(os.path.dirname(__file__), "wf_shared_parameters.txt")

PANEL_WIDTH_FT = 4.0
PANEL_HEIGHT_FT = 8.0
PANEL_AREA_SF = PANEL_WIDTH_FT * PANEL_HEIGHT_FT
AREA_TOL = 1e-6
LENGTH_TOL = 1e-6


BOM_PARAMETER_DEFS = (
    {
        "name": "WF_IsGenerated",
        "guid": "569bf159-1854-475a-b236-7af05f81fb5e",
        "spec": "yesno",
        "categories": (
            DB.BuiltInCategory.OST_StructuralFraming,
            DB.BuiltInCategory.OST_StructuralColumns,
        ),
        "group": "data",
        "hide_when_no_value": False,
        "description": "Tracks framing instances managed by the Wood Framing extension.",
    },
    {
        "name": "WF_HostKind",
        "guid": "d108d3b6-10cb-4348-a817-77209a0f7f10",
        "spec": "text",
        "categories": (
            DB.BuiltInCategory.OST_StructuralFraming,
            DB.BuiltInCategory.OST_StructuralColumns,
        ),
        "group": "data",
        "hide_when_no_value": True,
        "description": "Host category display name for grouped framing schedules.",
    },
    {
        "name": "WF_HostId",
        "guid": "72a28f2a-97de-4eb6-a345-de7322e04371",
        "spec": "text",
        "categories": (
            DB.BuiltInCategory.OST_StructuralFraming,
            DB.BuiltInCategory.OST_StructuralColumns,
        ),
        "group": "data",
        "hide_when_no_value": True,
        "description": "Stable host element id used to group framing schedule rows.",
    },
    {
        "name": "WF_HostLabel",
        "guid": "43b8d63e-5506-4447-9b97-10a48d5fd3ea",
        "spec": "text",
        "categories": (
            DB.BuiltInCategory.OST_StructuralFraming,
            DB.BuiltInCategory.OST_StructuralColumns,
        ),
        "group": "data",
        "hide_when_no_value": True,
        "description": "Preferred host label for schedules, using Mark when available.",
    },
    {
        "name": "WF_MemberRole",
        "guid": "3a57b3b9-a61f-40f7-9021-383fcbef2465",
        "spec": "text",
        "categories": (
            DB.BuiltInCategory.OST_StructuralFraming,
            DB.BuiltInCategory.OST_StructuralColumns,
        ),
        "group": "data",
        "hide_when_no_value": True,
        "description": "Normalized framing member role used by the BOM schedule.",
    },
    {
        "name": "WF_MemberLength",
        "guid": "54e36114-4a85-4112-9859-c86d302472ee",
        "spec": "length",
        "categories": (
            DB.BuiltInCategory.OST_StructuralFraming,
            DB.BuiltInCategory.OST_StructuralColumns,
        ),
        "group": "data",
        "hide_when_no_value": False,
        "description": "Length used for BOM totals.",
    },
)


SHEATHING_PARAMETER_DEFS = (
    {
        "name": "WF_SheathHostLabel",
        "guid": "3a01f455-98f0-4800-893d-ff2ea7e5fd22",
        "spec": "text",
        "categories": (
            DB.BuiltInCategory.OST_Walls,
            DB.BuiltInCategory.OST_Floors,
            DB.BuiltInCategory.OST_Ceilings,
            DB.BuiltInCategory.OST_Roofs,
        ),
        "group": "data",
        "hide_when_no_value": True,
        "description": "Preferred host label for the sheathing schedule.",
    },
    {
        "name": "WF_SheathFullSheets",
        "guid": "3412bf34-b6f9-40f7-92f8-5c6fdeaadf59",
        "spec": "integer",
        "categories": (
            DB.BuiltInCategory.OST_Walls,
            DB.BuiltInCategory.OST_Floors,
            DB.BuiltInCategory.OST_Ceilings,
            DB.BuiltInCategory.OST_Roofs,
        ),
        "group": "data",
        "hide_when_no_value": False,
        "description": "Count of full 4x8 sheathing sheets.",
    },
    {
        "name": "WF_SheathCutCount",
        "guid": "fe8357bf-0744-4121-befb-a07eed6a30fb",
        "spec": "integer",
        "categories": (
            DB.BuiltInCategory.OST_Walls,
            DB.BuiltInCategory.OST_Floors,
            DB.BuiltInCategory.OST_Ceilings,
            DB.BuiltInCategory.OST_Roofs,
        ),
        "group": "data",
        "hide_when_no_value": False,
        "description": "Count of partial sheathing panels.",
    },
    {
        "name": "WF_SheathCutArea",
        "guid": "a8f17aa4-a225-423b-82da-2c2e672f7e66",
        "spec": "area",
        "categories": (
            DB.BuiltInCategory.OST_Walls,
            DB.BuiltInCategory.OST_Floors,
            DB.BuiltInCategory.OST_Ceilings,
            DB.BuiltInCategory.OST_Roofs,
        ),
        "group": "data",
        "hide_when_no_value": False,
        "description": "Total reusable partial-sheet area.",
    },
    {
        "name": "WF_SheathCutSheetEq",
        "guid": "7ce6ae15-42a9-45ec-8806-a4dd77a3c087",
        "spec": "number",
        "categories": (
            DB.BuiltInCategory.OST_Walls,
            DB.BuiltInCategory.OST_Floors,
            DB.BuiltInCategory.OST_Ceilings,
            DB.BuiltInCategory.OST_Roofs,
        ),
        "group": "data",
        "hide_when_no_value": False,
        "description": "Partial-sheet total converted into 32 sf sheet equivalents.",
    },
    {
        "name": "WF_SheathTotalSheetEq",
        "guid": "57bba2f1-36d3-4398-a394-0fc6b80cada4",
        "spec": "number",
        "categories": (
            DB.BuiltInCategory.OST_Walls,
            DB.BuiltInCategory.OST_Floors,
            DB.BuiltInCategory.OST_Ceilings,
            DB.BuiltInCategory.OST_Roofs,
        ),
        "group": "data",
        "hide_when_no_value": False,
        "description": "Full sheets plus reusable partial-sheet equivalents.",
    },
    {
        "name": "WF_SheathCutSummary",
        "guid": "847a73d7-40e5-4474-b7c6-ca16f49e14da",
        "spec": "text",
        "categories": (
            DB.BuiltInCategory.OST_Walls,
            DB.BuiltInCategory.OST_Floors,
            DB.BuiltInCategory.OST_Ceilings,
            DB.BuiltInCategory.OST_Roofs,
        ),
        "group": "data",
        "hide_when_no_value": True,
        "description": "Compact cut-size summary for the sheathing schedule.",
    },
)


def ensure_bom_parameters(doc):
    """Ensure all BOM shared parameters exist and are bound."""
    _ensure_shared_parameters(doc, BOM_PARAMETER_DEFS)


def ensure_sheathing_parameters(doc):
    """Ensure all sheathing shared parameters exist and are bound."""
    _ensure_shared_parameters(doc, SHEATHING_PARAMETER_DEFS)


def apply_bom_metadata(instance, host_info, member_role, length_ft=None):
    """Write BOM metadata directly to a framing instance."""
    if instance is None or host_info is None:
        return

    doc = instance.Document
    host_kind = _display_host_kind(getattr(host_info, "kind", None))
    host_element = getattr(host_info, "element", None)
    host_id = getattr(host_info, "element_id", None)
    if host_element is None and host_id is not None:
        host_element = _get_element_by_id(doc, host_id)

    _write_bom_metadata(
        instance,
        True,
        host_kind,
        _element_id_string(host_id),
        _host_label(doc, host_element, host_kind, host_id),
        _member_role_label(member_role),
        length_ft if length_ft is not None else _element_length_ft(instance),
    )


def apply_bom_metadata_from_member(instance, host_info, member):
    """Write BOM metadata using a generated member descriptor."""
    if instance is None or host_info is None or member is None:
        return

    doc = instance.Document
    host_id = getattr(member, "host_id", None) or getattr(host_info, "element_id", None)
    host_element = _get_element_by_id(doc, host_id)
    host_kind = _display_host_kind(
        getattr(member, "host_kind", None) or getattr(host_info, "kind", None)
    )

    apply_bom_metadata(
        instance,
        _SimpleHostInfo(host_kind, host_id, host_element),
        getattr(member, "member_type", None),
    )


def backfill_bom_metadata(doc):
    """Backfill shared BOM parameters from tracking data when possible."""
    for element in _collect_elements(
        doc,
        (
            DB.BuiltInCategory.OST_StructuralFraming,
            DB.BuiltInCategory.OST_StructuralColumns,
        ),
    ):
        tracking = get_tracking_data(element)
        if tracking is not None:
            host_kind_key = tracking.get("kind")
            host_id_text = tracking.get("host")
            if host_kind_key and host_id_text:
                host_kind = _display_host_kind(host_kind_key)
                host_element = _get_element_by_id(doc, host_id_text)
                _write_bom_metadata(
                    element,
                    True,
                    host_kind,
                    str(host_id_text),
                    _host_label(doc, host_element, host_kind, host_id_text),
                    _member_role_label(tracking.get("member")),
                    _element_length_ft(element),
                )
                continue

        comments = _comments_text(element)
        if "WF_Generated" in comments:
            _clear_bom_metadata(element)


def clear_all_sheathing_metadata(doc):
    """Reset sheathing metadata so the schedule reflects the current run only."""
    for element in _collect_elements(
        doc,
        (
            DB.BuiltInCategory.OST_Walls,
            DB.BuiltInCategory.OST_Floors,
            DB.BuiltInCategory.OST_Ceilings,
            DB.BuiltInCategory.OST_Roofs,
        ),
    ):
        _clear_sheathing_metadata(element)


def calculate_sheathing_for_host(doc, element):
    """Calculate sheathing metrics for a supported host element."""
    cfg = FramingConfig()

    if isinstance(element, DB.Wall):
        host_info = analyze_wall_host(doc, element, cfg)
        if host_info is None:
            return None
        loops = _wall_loops(host_info)
        result = _panelize_loops(
            loops,
            0.0,
            host_info.length,
            0.0,
            max(host_info.start_height, host_info.end_height, host_info.height),
            PANEL_WIDTH_FT,
            PANEL_HEIGHT_FT,
        )
        return _finalize_sheathing_result(doc, element, result)

    if isinstance(element, DB.Floor):
        host_info = analyze_floor_host(doc, element, cfg)
        if host_info is None:
            return None
        result = _panelize_planar_host(host_info)
        return _finalize_sheathing_result(doc, element, result)

    if _is_ceiling(element):
        host_info = analyze_ceiling_host(doc, element, cfg)
        if host_info is None:
            return None
        result = _panelize_planar_host(host_info)
        return _finalize_sheathing_result(doc, element, result)

    if isinstance(element, DB.RoofBase):
        roof_info = analyze_roof_host(doc, element, cfg)
        if roof_info is None:
            return None
        aggregate = _empty_sheathing_result()
        for plane in getattr(roof_info, "planes", []) or []:
            plane_result = _panelize_planar_host(plane)
            _merge_sheathing_result(aggregate, plane_result)
        return _finalize_sheathing_result(doc, element, aggregate)

    return None


def stamp_sheathing_metadata(element, result):
    """Write sheathing metrics to a host element."""
    if element is None or result is None:
        return

    _set_shared_text(element, "WF_SheathHostLabel", result.get("host_label"))
    _set_shared_int(element, "WF_SheathFullSheets", result.get("full_sheets", 0))
    _set_shared_int(element, "WF_SheathCutCount", result.get("cut_count", 0))
    _set_shared_double(element, "WF_SheathCutArea", result.get("cut_area", 0.0))
    _set_shared_double(element, "WF_SheathCutSheetEq", result.get("cut_sheet_eq", 0.0))
    _set_shared_double(element, "WF_SheathTotalSheetEq", result.get("total_sheet_eq", 0.0))
    _set_shared_text(element, "WF_SheathCutSummary", result.get("cut_summary"))


def create_or_update_bom_schedule(doc):
    """Create or refresh the BOM schedule."""
    ensure_bom_parameters(doc)
    backfill_bom_metadata(doc)

    schedule = _find_schedule_by_name(doc, BOM_SCHEDULE_NAME)
    if schedule is None:
        schedule = DB.ViewSchedule.CreateSchedule(
            doc,
            DB.ElementId.InvalidElementId,
            DB.ElementId.InvalidElementId,
        )
    schedule.Name = BOM_SCHEDULE_NAME

    definition = schedule.Definition
    definition.ClearFields()
    definition.ClearFilters()
    definition.ClearSortGroupFields()
    definition.IsItemized = False
    definition.ShowGrandTotal = True
    definition.ShowGrandTotalCount = True
    definition.ShowGrandTotalTitle = True

    fields = OrderedDict()
    fields["generated"] = _add_shared_field(definition, doc, "WF_IsGenerated")
    fields["host_id"] = _add_shared_field(definition, doc, "WF_HostId")
    fields["host_kind"] = _add_shared_field(definition, doc, "WF_HostKind")
    fields["host_label"] = _add_shared_field(definition, doc, "WF_HostLabel")
    fields["member_role"] = _add_shared_field(definition, doc, "WF_MemberRole")
    fields["family_type"] = definition.AddField(
        DB.ScheduleFieldType.Instance,
        DB.ElementId(DB.BuiltInParameter.ELEM_FAMILY_AND_TYPE_PARAM),
    )
    fields["count"] = definition.AddField(DB.ScheduleFieldType.Count)
    fields["member_length"] = _add_shared_field(definition, doc, "WF_MemberLength")

    fields["generated"].IsHidden = True
    fields["host_id"].IsHidden = True
    fields["host_kind"].ColumnHeading = "Host Kind"
    fields["host_label"].ColumnHeading = "Host"
    fields["member_role"].ColumnHeading = "Member Role"
    fields["family_type"].ColumnHeading = "Family / Type"
    fields["count"].ColumnHeading = "Count"
    fields["member_length"].ColumnHeading = "Total Length"
    if fields["member_length"].CanTotal():
        fields["member_length"].DisplayType = DB.ScheduleFieldDisplayType.Totals

    for key in ("host_id", "host_label", "member_role", "family_type"):
        definition.AddSortGroupField(
            DB.ScheduleSortGroupField(fields[key].FieldId, DB.ScheduleSortOrder.Ascending)
        )

    definition.AddFilter(
        DB.ScheduleFilter(fields["generated"].FieldId, DB.ScheduleFilterType.Equal, 1)
    )
    definition.AddFilter(
        DB.ScheduleFilter(fields["host_id"].FieldId, DB.ScheduleFilterType.HasValue)
    )

    return schedule


def create_or_update_sheathing_schedule(doc):
    """Create or refresh the sheathing schedule."""
    ensure_sheathing_parameters(doc)

    schedule = _find_schedule_by_name(doc, SHEATHING_SCHEDULE_NAME)
    if schedule is None:
        schedule = DB.ViewSchedule.CreateSchedule(
            doc,
            DB.ElementId.InvalidElementId,
            DB.ElementId.InvalidElementId,
        )
    schedule.Name = SHEATHING_SCHEDULE_NAME

    definition = schedule.Definition
    definition.ClearFields()
    definition.ClearFilters()
    definition.ClearSortGroupFields()
    definition.IsItemized = True
    definition.ShowGrandTotal = True
    definition.ShowGrandTotalCount = True
    definition.ShowGrandTotalTitle = True

    fields = OrderedDict()
    fields["host_label"] = _add_shared_field(definition, doc, "WF_SheathHostLabel")
    fields["full_sheets"] = _add_shared_field(definition, doc, "WF_SheathFullSheets")
    fields["cut_count"] = _add_shared_field(definition, doc, "WF_SheathCutCount")
    fields["cut_area"] = _add_shared_field(definition, doc, "WF_SheathCutArea")
    fields["cut_sheet_eq"] = _add_shared_field(definition, doc, "WF_SheathCutSheetEq")
    fields["total_sheet_eq"] = _add_shared_field(definition, doc, "WF_SheathTotalSheetEq")
    fields["cut_summary"] = _add_shared_field(definition, doc, "WF_SheathCutSummary")

    fields["host_label"].ColumnHeading = "Host"
    fields["full_sheets"].ColumnHeading = "Full Sheets"
    fields["cut_count"].ColumnHeading = "Cut Count"
    fields["cut_area"].ColumnHeading = "Cut Area"
    fields["cut_sheet_eq"].ColumnHeading = "Cut Sheet Eq"
    fields["total_sheet_eq"].ColumnHeading = "Total Sheet Eq"
    fields["cut_summary"].ColumnHeading = "Cut Summary"
    for key in ("full_sheets", "cut_count", "cut_area", "cut_sheet_eq", "total_sheet_eq"):
        try:
            if fields[key].CanTotal():
                fields[key].DisplayType = DB.ScheduleFieldDisplayType.Totals
        except Exception:
            pass

    definition.AddSortGroupField(
        DB.ScheduleSortGroupField(fields["host_label"].FieldId, DB.ScheduleSortOrder.Ascending)
    )
    definition.AddFilter(
        DB.ScheduleFilter(fields["host_label"].FieldId, DB.ScheduleFilterType.HasValue)
    )

    return schedule


def activate_schedule(schedule):
    """Try to show the resulting schedule to the user."""
    if schedule is None:
        return
    try:
        revit.uidoc.ActiveView = schedule
    except Exception:
        pass


class _SimpleHostInfo(object):
    """Small host shim used when only host metadata is available."""

    def __init__(self, kind, element_id, element):
        self.kind = kind
        self.element_id = element_id
        self.element = element


def _ensure_shared_parameters(doc, definitions):
    _ensure_shared_parameter_file()

    app = doc.Application
    previous_path = None
    try:
        previous_path = app.SharedParametersFilename
    except Exception:
        previous_path = None

    try:
        app.SharedParametersFilename = SHARED_PARAM_FILE
        definition_file = None
        try:
            definition_file = app.OpenSharedParameterFile()
        except Exception:
            definition_file = None
        if definition_file is None:
            # Recover from corrupted/mixed-encoding files (readParamDatabase).
            _write_shared_parameter_file()
            app.SharedParametersFilename = SHARED_PARAM_FILE
            definition_file = app.OpenSharedParameterFile()
        if definition_file is None:
            raise Exception("Could not open shared parameter file.")

        group = _get_or_create_group(definition_file.Groups, SHARED_PARAM_GROUP_NAME)
        for param_def in definitions:
            definition = _find_definition(definition_file, param_def["name"])
            if definition is None:
                definition = _create_definition(group, param_def)
            _bind_definition(doc, definition, param_def["categories"], param_def["group"])
    finally:
        try:
            app.SharedParametersFilename = previous_path or ""
        except Exception:
            pass


def _ensure_shared_parameter_file():
    if not os.path.exists(SHARED_PARAM_FILE):
        _write_shared_parameter_file()
        return

    # Rebuild if file has mixed/binary content (common after partial writes).
    try:
        with open(SHARED_PARAM_FILE, "rb") as stream:
            raw = stream.read()
    except Exception:
        _write_shared_parameter_file()
        return

    if b"\x00" in raw:
        _write_shared_parameter_file()
        return

    text = None
    for encoding in ("utf-8", "utf-8-sig", "cp1252"):
        try:
            text = raw.decode(encoding)
            break
        except Exception:
            continue
    if text is None:
        _write_shared_parameter_file()
        return

    required_tokens = [
        "*META\tVERSION\tMINVERSION",
        "*GROUP\tID\tNAME",
        "*PARAM\tGUID\tNAME\tDATATYPE\tDATACATEGORY\tGROUP\tVISIBLE\tDESCRIPTION\tUSERMODIFIABLE\tHIDEWHENNOVALUE",
        "WF_IsGenerated",
        "WF_SheathHostLabel",
    ]
    for token in required_tokens:
        if token not in text:
            _write_shared_parameter_file()
            return


def _write_shared_parameter_file():
    lines = [
        "# This is a Revit shared parameter file.",
        "# Do not edit manually.",
        "*META\tVERSION\tMINVERSION",
        "META\t2\t1",
        "*GROUP\tID\tNAME",
        "GROUP\t1\t{0}".format(SHARED_PARAM_GROUP_NAME),
        "*PARAM\tGUID\tNAME\tDATATYPE\tDATACATEGORY\tGROUP\tVISIBLE\tDESCRIPTION\tUSERMODIFIABLE\tHIDEWHENNOVALUE",
    ]

    for param_def in _all_shared_parameter_defs():
        lines.append(
            "PARAM\t{0}\t{1}\t{2}\t\t1\t1\t{3}\t1\t{4}".format(
                param_def["guid"],
                param_def["name"],
                _shared_param_datatype(param_def["spec"]),
                _shared_param_text(param_def.get("description")),
                "1" if param_def.get("hide_when_no_value") else "0",
            )
        )

    with open(SHARED_PARAM_FILE, "wb") as stream:
        stream.write(("\n".join(lines) + "\n").encode("utf-8"))


def _all_shared_parameter_defs():
    merged = []
    seen = set()
    for param_def in list(BOM_PARAMETER_DEFS) + list(SHEATHING_PARAMETER_DEFS):
        name = param_def.get("name")
        if not name or name in seen:
            continue
        seen.add(name)
        merged.append(param_def)
    return merged


def _shared_param_datatype(spec_name):
    mapping = {
        "yesno": "YESNO",
        "text": "TEXT",
        "length": "LENGTH",
        "area": "AREA",
        "number": "NUMBER",
        "integer": "INTEGER",
    }
    return mapping.get(spec_name, "TEXT")


def _shared_param_text(value):
    if value is None:
        return ""
    text = str(value)
    return text.replace("\t", " ").replace("\r", " ").replace("\n", " ")


def _get_or_create_group(groups, group_name):
    group = _collection_get(groups, group_name)
    if group is not None:
        return group
    return groups.Create(group_name)


def _find_definition(definition_file, name):
    for group in definition_file.Groups:
        definition = _collection_get(group.Definitions, name)
        if definition is not None:
            return definition
    return None


def _create_definition(group, param_def):
    from System import Guid

    data_type = _spec_type_id(param_def["spec"])
    options = DB.ExternalDefinitionCreationOptions(param_def["name"], data_type)
    try:
        options.GUID = Guid(param_def["guid"])
    except Exception:
        pass
    try:
        options.Description = param_def.get("description") or ""
    except Exception:
        pass
    try:
        options.UserModifiable = True
    except Exception:
        pass
    try:
        options.Visible = True
    except Exception:
        pass
    if param_def.get("hide_when_no_value"):
        try:
            options.HideWhenNoValue = True
        except Exception:
            pass
    return group.Definitions.Create(options)


def _spec_type_id(spec_name):
    if hasattr(DB, "SpecTypeId"):
        if spec_name == "yesno":
            return DB.SpecTypeId.Boolean.YesNo
        if spec_name == "text":
            return DB.SpecTypeId.String.Text
        if spec_name == "length":
            return getattr(DB.SpecTypeId, "Length", getattr(DB.SpecTypeId, "Distance"))
        if spec_name == "area":
            return DB.SpecTypeId.Area
        if spec_name == "number":
            return DB.SpecTypeId.Number
        if spec_name == "integer":
            int_class = getattr(DB.SpecTypeId, "Int", None)
            if int_class is not None:
                int_spec = getattr(int_class, "Integer", None)
                if int_spec is not None:
                    return int_spec
            return DB.SpecTypeId.Number
    return getattr(DB.ParameterType, "Text", None)


def _bind_definition(doc, definition, categories, group_name):
    binding_map = doc.ParameterBindings
    existing_def, existing_binding = _find_binding(binding_map, definition.Name)

    category_set = doc.Application.Create.NewCategorySet()
    inserted_ids = set()

    for category_id in categories:
        category = _category_from_builtin(doc, category_id)
        _insert_category_once(category_set, inserted_ids, category)

    if existing_binding is not None:
        try:
            for category in existing_binding.Categories:
                _insert_category_once(category_set, inserted_ids, category)
        except Exception:
            pass

    binding = doc.Application.Create.NewInstanceBinding(category_set)
    group_type_id = _group_type_id(group_name)

    target_def = existing_def or definition
    inserted = False
    try:
        if group_type_id is not None:
            inserted = binding_map.Insert(target_def, binding, group_type_id)
        else:
            inserted = binding_map.Insert(target_def, binding)
    except Exception:
        inserted = False

    if inserted:
        return

    try:
        if group_type_id is not None:
            binding_map.ReInsert(target_def, binding, group_type_id)
        else:
            binding_map.ReInsert(target_def, binding)
    except Exception:
        pass


def _category_from_builtin(doc, category_id):
    try:
        return DB.Category.GetCategory(doc, category_id)
    except Exception:
        return None


def _insert_category_once(category_set, inserted_ids, category):
    if category is None:
        return
    category_key = _category_id_int(getattr(category, "Id", None))
    if category_key in inserted_ids:
        return
    try:
        category_set.Insert(category)
        inserted_ids.add(category_key)
    except Exception:
        pass


def _find_binding(binding_map, definition_name):
    iterator = binding_map.ForwardIterator()
    iterator.Reset()
    while iterator.MoveNext():
        definition = iterator.Key
        if getattr(definition, "Name", None) == definition_name:
            return definition, iterator.Current
    return None, None


def _group_type_id(group_name):
    if not hasattr(DB, "GroupTypeId"):
        return None
    if group_name == "structural":
        return getattr(DB.GroupTypeId, "Structural", None)
    return getattr(DB.GroupTypeId, "Data", None)


def _find_schedule_by_name(doc, name):
    collector = DB.FilteredElementCollector(doc).OfClass(DB.ViewSchedule)
    for schedule in collector:
        if getattr(schedule, "Name", None) == name:
            return schedule
    return None


def _add_shared_field(definition, doc, parameter_name):
    parameter_id = _shared_parameter_element_id(doc, parameter_name)
    return definition.AddField(DB.ScheduleFieldType.Instance, parameter_id)


def _shared_parameter_element_id(doc, parameter_name):
    collector = DB.FilteredElementCollector(doc).OfClass(DB.SharedParameterElement)
    for param_element in collector:
        definition = param_element.GetDefinition()
        if definition is not None and definition.Name == parameter_name:
            return param_element.Id
    raise Exception("Shared parameter not found: {0}".format(parameter_name))


def _write_bom_metadata(element, is_generated, host_kind, host_id, host_label,
                        member_role, member_length_ft):
    _set_shared_int(element, "WF_IsGenerated", 1 if is_generated else 0)
    _set_shared_text(element, "WF_HostKind", host_kind)
    _set_shared_text(element, "WF_HostId", host_id)
    _set_shared_text(element, "WF_HostLabel", host_label)
    _set_shared_text(element, "WF_MemberRole", member_role)
    _set_shared_double(element, "WF_MemberLength", member_length_ft or 0.0)


def _clear_bom_metadata(element):
    _set_shared_int(element, "WF_IsGenerated", 0)
    _set_shared_text(element, "WF_HostKind", None)
    _set_shared_text(element, "WF_HostId", None)
    _set_shared_text(element, "WF_HostLabel", None)
    _set_shared_text(element, "WF_MemberRole", None)
    _set_shared_double(element, "WF_MemberLength", 0.0)


def _clear_sheathing_metadata(element):
    _set_shared_text(element, "WF_SheathHostLabel", None)
    _set_shared_int(element, "WF_SheathFullSheets", 0)
    _set_shared_int(element, "WF_SheathCutCount", 0)
    _set_shared_double(element, "WF_SheathCutArea", 0.0)
    _set_shared_double(element, "WF_SheathCutSheetEq", 0.0)
    _set_shared_double(element, "WF_SheathTotalSheetEq", 0.0)
    _set_shared_text(element, "WF_SheathCutSummary", None)


def _set_shared_text(element, name, value):
    parameter = element.LookupParameter(name)
    if parameter is None or parameter.IsReadOnly:
        return

    text = value if value not in (None, "") else None
    if text is None:
        try:
            parameter.ClearValue()
            return
        except Exception:
            pass
        try:
            parameter.Set("")
        except Exception:
            pass
        return

    try:
        parameter.Set(str(text))
    except Exception:
        pass


def _set_shared_int(element, name, value):
    parameter = element.LookupParameter(name)
    if parameter is None or parameter.IsReadOnly:
        return
    try:
        parameter.Set(int(value or 0))
    except Exception:
        pass


def _set_shared_double(element, name, value):
    parameter = element.LookupParameter(name)
    if parameter is None or parameter.IsReadOnly:
        return
    try:
        parameter.Set(float(value or 0.0))
    except Exception:
        pass


def _collect_elements(doc, categories):
    seen = set()
    elements = []
    for category in categories:
        collector = (
            DB.FilteredElementCollector(doc)
            .OfCategory(category)
            .WhereElementIsNotElementType()
        )
        for element in collector:
            element_id = _element_id_string(element.Id)
            if element_id in seen:
                continue
            seen.add(element_id)
            elements.append(element)
    return elements


def _comments_text(element):
    parameter = element.get_Parameter(DB.BuiltInParameter.ALL_MODEL_INSTANCE_COMMENTS)
    if parameter is None:
        return ""
    try:
        return parameter.AsString() or ""
    except Exception:
        return ""


def _host_label(doc, element, host_kind, host_id):
    mark = None
    if element is not None:
        parameter = element.get_Parameter(DB.BuiltInParameter.ALL_MODEL_MARK)
        if parameter is None:
            try:
                parameter = element.LookupParameter("Mark")
            except Exception:
                parameter = None
        if parameter is not None:
            try:
                mark = parameter.AsString()
            except Exception:
                mark = None
    if mark:
        return mark
    return "{0} {1}".format(host_kind or "Host", _element_id_string(host_id))


def _display_host_kind(kind):
    raw = (kind or "host").strip()
    if not raw:
        raw = "host"
    return raw.replace("_", " ").replace("-", " ").title()


def _member_role_label(member_role):
    raw = (member_role or "generated").strip()
    if not raw:
        raw = "generated"
    words = raw.replace("_", " ").replace("-", " ").split()
    return " ".join([word[:1].upper() + word[1:].lower() for word in words])


def _element_length_ft(element):
    for parameter_id in (
        DB.BuiltInParameter.STRUCTURAL_FRAME_CUT_LENGTH,
        DB.BuiltInParameter.INSTANCE_LENGTH_PARAM,
    ):
        parameter = element.get_Parameter(parameter_id)
        if parameter is not None and parameter.HasValue:
            try:
                value = parameter.AsDouble()
            except Exception:
                value = 0.0
            if value > 0.0:
                return value

    location = getattr(element, "Location", None)
    curve = getattr(location, "Curve", None)
    if curve is not None:
        try:
            return curve.Length
        except Exception:
            pass

    box = element.get_BoundingBox(None)
    if box is not None:
        try:
            return max(
                abs(box.Max.X - box.Min.X),
                abs(box.Max.Y - box.Min.Y),
                abs(box.Max.Z - box.Min.Z),
            )
        except Exception:
            pass

    return 0.0


def _wall_loops(host_info):
    outer = [
        (0.0, 0.0),
        (host_info.length, 0.0),
        (host_info.length, max(0.0, host_info.end_height)),
        (0.0, max(0.0, host_info.start_height)),
    ]
    loops = [outer]

    for opening in getattr(host_info, "openings", []) or []:
        left = max(0.0, min(host_info.length, opening.left_edge))
        right = max(0.0, min(host_info.length, opening.right_edge))
        sill = max(0.0, opening.sill_height)
        head = max(sill, opening.head_height)
        if right - left < LENGTH_TOL or head - sill < LENGTH_TOL:
            continue
        loops.append(
            [
                (left, sill),
                (right, sill),
                (right, head),
                (left, head),
            ]
        )

    return loops


def _panelize_planar_host(host_info):
    min_x, max_x, min_y, max_y = host_info.bounds
    return _panelize_loops(
        host_info.boundary_loops_local,
        min_x,
        max_x,
        min_y,
        max_y,
        PANEL_WIDTH_FT,
        PANEL_HEIGHT_FT,
    )


def _empty_sheathing_result():
    return {
        "full_sheets": 0,
        "cut_count": 0,
        "cut_area": 0.0,
        "cut_sheet_eq": 0.0,
        "total_sheet_eq": 0.0,
        "cut_summary_counts": OrderedDict(),
        "cut_summary": None,
    }


def _merge_sheathing_result(target, result):
    target["full_sheets"] += result.get("full_sheets", 0)
    target["cut_count"] += result.get("cut_count", 0)
    target["cut_area"] += result.get("cut_area", 0.0)
    for label, count in (result.get("cut_summary_counts") or {}).items():
        target["cut_summary_counts"][label] = target["cut_summary_counts"].get(label, 0) + count


def _finalize_sheathing_result(doc, element, result):
    final = _empty_sheathing_result()
    _merge_sheathing_result(final, result or {})
    final["cut_sheet_eq"] = final["cut_area"] / PANEL_AREA_SF
    final["total_sheet_eq"] = final["full_sheets"] + final["cut_sheet_eq"]
    final["host_label"] = _host_label(
        doc,
        element,
        _display_host_kind(_host_kind_key(element)),
        element.Id,
    )
    final["cut_summary"] = _cut_summary_text(final["cut_summary_counts"])
    return final


def _panelize_loops(loops_local, min_x, max_x, min_y, max_y, panel_w, panel_h):
    result = _empty_sheathing_result()
    if not loops_local:
        return result

    x = min_x
    while x < max_x - LENGTH_TOL:
        x1 = min(x + panel_w, max_x)
        y = min_y
        while y < max_y - LENGTH_TOL:
            y1 = min(y + panel_h, max_y)
            area, bbox = _clipped_panel_area_and_bbox(loops_local, x, x1, y, y1)
            if area <= AREA_TOL:
                y = y1
                continue

            panel_area = (x1 - x) * (y1 - y)
            is_full = (
                abs((x1 - x) - panel_w) <= LENGTH_TOL
                and abs((y1 - y) - panel_h) <= LENGTH_TOL
                and abs(area - panel_area) <= AREA_TOL
            )
            if is_full:
                result["full_sheets"] += 1
            else:
                result["cut_count"] += 1
                result["cut_area"] += area
                label = _cut_label_from_bbox(bbox)
                result["cut_summary_counts"][label] = result["cut_summary_counts"].get(label, 0) + 1
            y = y1
        x = x1

    return result


def _clipped_panel_area_and_bbox(loops_local, x0, x1, y0, y1):
    if not loops_local:
        return 0.0, None

    outer_index = _outer_loop_index(loops_local)
    total_area = 0.0
    bbox = None

    for index, loop in enumerate(loops_local):
        clipped = _clip_polygon_to_rect(loop, x0, x1, y0, y1)
        if len(clipped) < 3:
            continue

        area = abs(_polygon_area(clipped))
        if area <= AREA_TOL:
            continue

        if index == outer_index:
            total_area += area
            bbox = _merge_bbox(bbox, _polygon_bbox(clipped))
        else:
            total_area -= area

    return max(0.0, total_area), bbox


def _outer_loop_index(loops_local):
    best_index = 0
    best_area = -1.0
    for index, loop in enumerate(loops_local):
        area = abs(_polygon_area(loop))
        if area > best_area:
            best_area = area
            best_index = index
    return best_index


def _clip_polygon_to_rect(polygon, x_min, x_max, y_min, y_max):
    clipped = list(polygon)
    for edge_name, edge_value in (
        ("left", x_min),
        ("right", x_max),
        ("bottom", y_min),
        ("top", y_max),
    ):
        clipped = _clip_polygon_edge(clipped, edge_name, edge_value)
        if not clipped:
            break
    return clipped


def _clip_polygon_edge(points, edge_name, edge_value):
    if not points:
        return []

    result = []
    previous = points[-1]
    previous_inside = _point_inside(previous, edge_name, edge_value)
    for current in points:
        current_inside = _point_inside(current, edge_name, edge_value)
        if current_inside:
            if not previous_inside:
                result.append(_edge_intersection(previous, current, edge_name, edge_value))
            result.append(current)
        elif previous_inside:
            result.append(_edge_intersection(previous, current, edge_name, edge_value))
        previous = current
        previous_inside = current_inside
    return result


def _point_inside(point, edge_name, edge_value):
    x, y = point
    if edge_name == "left":
        return x >= edge_value - LENGTH_TOL
    if edge_name == "right":
        return x <= edge_value + LENGTH_TOL
    if edge_name == "bottom":
        return y >= edge_value - LENGTH_TOL
    return y <= edge_value + LENGTH_TOL


def _edge_intersection(start, end, edge_name, edge_value):
    x0, y0 = start
    x1, y1 = end

    if edge_name in ("left", "right"):
        if abs(x1 - x0) <= LENGTH_TOL:
            return (edge_value, y0)
        t = (edge_value - x0) / float(x1 - x0)
        return (edge_value, y0 + (y1 - y0) * t)

    if abs(y1 - y0) <= LENGTH_TOL:
        return (x0, edge_value)
    t = (edge_value - y0) / float(y1 - y0)
    return (x0 + (x1 - x0) * t, edge_value)


def _polygon_area(points):
    area = 0.0
    count = len(points)
    for index in range(count):
        x0, y0 = points[index]
        x1, y1 = points[(index + 1) % count]
        area += (x0 * y1) - (x1 * y0)
    return area * 0.5


def _polygon_bbox(points):
    xs = [point[0] for point in points]
    ys = [point[1] for point in points]
    return (min(xs), max(xs), min(ys), max(ys))


def _merge_bbox(first, second):
    if first is None:
        return second
    if second is None:
        return first
    return (
        min(first[0], second[0]),
        max(first[1], second[1]),
        min(first[2], second[2]),
        max(first[3], second[3]),
    )


def _cut_label_from_bbox(bbox):
    if bbox is None:
        return "partial"
    width = max(0.0, bbox[1] - bbox[0])
    height = max(0.0, bbox[3] - bbox[2])
    return "{0} x {1}".format(_format_feet_inches(width), _format_feet_inches(height))


def _cut_summary_text(summary_counts):
    if not summary_counts:
        return None
    parts = []
    for label, count in sorted(summary_counts.items()):
        parts.append("{0} x {1}".format(label, count))
    return "; ".join(parts)


def _format_feet_inches(length_ft):
    inches_total = max(0.0, length_ft) * 12.0
    rounded_inches = round(inches_total * 2.0) / 2.0
    feet = int(rounded_inches // 12.0)
    inches = rounded_inches - (feet * 12.0)
    if abs(inches - round(inches)) <= 1e-9:
        inch_text = str(int(round(inches)))
    else:
        inch_text = "{0:.1f}".format(inches).rstrip("0").rstrip(".")
    return "{0}'-{1}\"".format(feet, inch_text)


def _host_kind_key(element):
    if isinstance(element, DB.Wall):
        return "wall"
    if isinstance(element, DB.Floor):
        return "floor"
    if _is_ceiling(element):
        return "ceiling"
    if isinstance(element, DB.RoofBase):
        return "roof"
    return "host"


def _is_ceiling(element):
    try:
        category = getattr(element, "Category", None)
        if category is not None:
            return _category_id_int(category.Id) == int(DB.BuiltInCategory.OST_Ceilings)
    except Exception:
        pass
    if hasattr(DB, "Ceiling"):
        try:
            return isinstance(element, DB.Ceiling)
        except Exception:
            pass
    return False


def _get_element_by_id(doc, raw_id):
    if raw_id is None:
        return None
    try:
        if hasattr(raw_id, "IntegerValue") or hasattr(raw_id, "Value"):
            return doc.GetElement(raw_id)
    except Exception:
        pass
    try:
        return doc.GetElement(DB.ElementId(int(raw_id)))
    except Exception:
        return None


def _element_id_string(raw_id):
    if raw_id is None:
        return ""
    if isinstance(raw_id, str):
        return raw_id
    return str(_category_id_int(raw_id))


def _category_id_int(raw_id):
    if raw_id is None:
        return -1
    if isinstance(raw_id, int):
        return raw_id
    return getattr(raw_id, "IntegerValue", getattr(raw_id, "Value", -1))


def _collection_get(collection, name):
    try:
        item = collection.get_Item(name)
        if item is not None:
            return item
    except Exception:
        pass
    try:
        return collection[name]
    except Exception:
        pass
    for item in collection:
        if getattr(item, "Name", None) == name:
            return item
    return None
