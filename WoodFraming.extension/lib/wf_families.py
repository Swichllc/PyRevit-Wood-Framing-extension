# -*- coding: utf-8 -*-
"""Family discovery and loading utilities.

Finds structural framing families in the current Revit project
and provides helpers to load families from external RFA files.
"""


def _get_name(element):
    """Get Element.Name safely in IronPython.

    IronPython 2.7 cannot always resolve the .NET Element.Name
    property via dynamic dispatch, so we use the descriptor directly.
    """
    from Autodesk.Revit.DB import Element
    return Element.Name.__get__(element)


def get_structural_framing_families(doc):
    """Collect all structural framing family symbols from the project.

    Returns:
        dict: {family_name: [type_name, ...]} sorted alphabetically.
    """
    from Autodesk.Revit.DB import (
        FilteredElementCollector,
        BuiltInCategory,
        FamilySymbol,
    )

    collector = (
        FilteredElementCollector(doc)
        .OfCategory(BuiltInCategory.OST_StructuralFraming)
        .OfClass(FamilySymbol)
    )

    families = {}
    for symbol in collector:
        fam_name = _get_name(symbol.Family)
        type_name = _get_name(symbol)
        if fam_name not in families:
            families[fam_name] = []
        if type_name not in families[fam_name]:
            families[fam_name].append(type_name)

    # Sort type names within each family
    for fam_name in families:
        families[fam_name].sort()

    return families


def get_structural_column_families(doc):
    """Collect all structural column family symbols from the project.

    Returns:
        dict: {family_name: [type_name, ...]}
    """
    from Autodesk.Revit.DB import (
        FilteredElementCollector,
        BuiltInCategory,
        FamilySymbol,
    )

    collector = (
        FilteredElementCollector(doc)
        .OfCategory(BuiltInCategory.OST_StructuralColumns)
        .OfClass(FamilySymbol)
    )

    families = {}
    for symbol in collector:
        fam_name = _get_name(symbol.Family)
        type_name = _get_name(symbol)
        if fam_name not in families:
            families[fam_name] = []
        if type_name not in families[fam_name]:
            families[fam_name].append(type_name)

    for fam_name in families:
        families[fam_name].sort()

    return families


def find_family_symbol(doc, family_name, type_name, category_bic=None):
    """Find a specific FamilySymbol by family name and type name.

    Args:
        doc: Revit Document
        family_name: str
        type_name: str
        category_bic: optional BuiltInCategory to narrow the search

    Returns:
        FamilySymbol or None
    """
    from Autodesk.Revit.DB import FilteredElementCollector, FamilySymbol

    collector = FilteredElementCollector(doc).OfClass(FamilySymbol)
    if category_bic is not None:
        collector = collector.OfCategory(category_bic)

    for symbol in collector:
        if _get_name(symbol.Family) == family_name and _get_name(symbol) == type_name:
            return symbol
    return None


def activate_symbol(doc, symbol):
    """Ensure a FamilySymbol is activated for placement.

    Args:
        doc: Revit Document (must be inside a transaction)
        symbol: FamilySymbol
    """
    if not symbol.IsActive:
        symbol.Activate()
        doc.Regenerate()


def load_family(doc, rfa_path):
    """Load a Revit family from an RFA file into the project.

    Args:
        doc: Revit Document (must be inside a transaction)
        rfa_path: full file path to the .rfa file

    Returns:
        Family object if successful, None otherwise.
    """
    import os
    if not os.path.exists(rfa_path):
        return None

    from Autodesk.Revit.DB import Family
    import clr

    # doc.LoadFamily returns (bool, Family) via out parameter
    # IronPython handles this as a tuple return
    result = clr.Reference[Family]()
    success = doc.LoadFamily(rfa_path, result)
    if success:
        return result.Value
    return None


def get_family_type_names(doc, family_name):
    """Get all type names for a given family name.

    Args:
        doc: Revit Document
        family_name: str

    Returns:
        list of type name strings
    """
    from Autodesk.Revit.DB import FilteredElementCollector, FamilySymbol

    types = []
    collector = FilteredElementCollector(doc).OfClass(FamilySymbol)
    for symbol in collector:
        if _get_name(symbol.Family) == family_name:
            types.append(_get_name(symbol))
    types.sort()
    return types


def get_available_types_flat(doc):
    """Get a flat list of 'FamilyName : TypeName' strings for UI display.

    Searches structural framing families only.

    Returns:
        list of str like ['Dimension Lumber : 2x4', ...]
    """
    from Autodesk.Revit.DB import (
        FilteredElementCollector,
        BuiltInCategory,
        FamilySymbol,
    )

    collector = (
        FilteredElementCollector(doc)
        .OfCategory(BuiltInCategory.OST_StructuralFraming)
        .OfClass(FamilySymbol)
    )

    items = []
    for symbol in collector:
        label = "{0} : {1}".format(_get_name(symbol.Family), _get_name(symbol))
        if label not in items:
            items.append(label)
    items.sort()
    return items


def get_column_types_flat(doc):
    """Get a flat list of 'FamilyName : TypeName' for structural columns.

    Returns:
        list of str like ['Wood Column : 6x6', ...]
    """
    from Autodesk.Revit.DB import (
        FilteredElementCollector,
        BuiltInCategory,
        FamilySymbol,
    )

    collector = (
        FilteredElementCollector(doc)
        .OfCategory(BuiltInCategory.OST_StructuralColumns)
        .OfClass(FamilySymbol)
    )

    items = []
    for symbol in collector:
        label = "{0} : {1}".format(_get_name(symbol.Family), _get_name(symbol))
        if label not in items:
            items.append(label)
    items.sort()
    return items


def parse_family_type_label(label):
    """Split a 'FamilyName : TypeName' label into (family, type) tuple."""
    parts = label.split(" : ", 1)
    if len(parts) == 2:
        return parts[0].strip(), parts[1].strip()
    return label.strip(), ""
