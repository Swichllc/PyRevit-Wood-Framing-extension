# -*- coding: utf-8 -*-
"""Shared placement service for all generated framing members.

Vertical members (studs, king studs, jack studs, cripples) are placed
with StructuralType.Column — Revit's intended type for vertical framing.

Horizontal members (plates, headers, sills) are placed with
StructuralType.Beam — Revit's intended type for horizontal framing.

Cross-section rotation (STRUCTURAL_BEND_DIR_ANGLE) is set as an absolute
value computed by the framing engine for each member.
"""

from wf_geometry import inches_to_feet
from wf_families import activate_symbol, find_family_symbol
from wf_tracking import tag_instance


MIN_MEMBER_LENGTH = inches_to_feet(1.0)


class BaseFramingEngine(object):
    """Shared placement and family resolution for framing engines."""

    def __init__(self, doc, config):
        self.doc = doc
        self.config = config

    # ------------------------------------------------------------------
    #  Placement
    # ------------------------------------------------------------------

    def place_members(self, members, host_info):
        """Place calculated framing members in Revit."""
        from Autodesk.Revit.DB import Line
        from Autodesk.Revit.DB.Structure import StructuralType

        try:
            from Autodesk.Revit.DB.Structure import StructuralFramingUtils
        except Exception:
            StructuralFramingUtils = None

        level = self.doc.GetElement(host_info.level_id)
        placed = []
        self._last_placed_pairs = []

        for member in members:
            symbol = self._resolve_symbol(member)
            if symbol is None:
                continue

            activate_symbol(self.doc, symbol)

            start = member.start_point
            end = member.end_point
            if self._curve_length(start, end) < MIN_MEMBER_LENGTH:
                continue

            # Vertical members → Column; horizontal → Beam.
            stype = (StructuralType.Column
                     if getattr(member, "is_column", False)
                     else StructuralType.Beam)

            is_col = getattr(member, "is_column", False)
            if is_col and self._is_vertical_member(member):
                instance = self._place_column_member(member, symbol, level)
            else:
                try:
                    line = Line.CreateBound(start, end)
                    instance = self.doc.Create.NewFamilyInstance(
                        line, symbol, level, stype
                    )
                except Exception:
                    instance = None

            if instance is None:
                continue

            # Center cross-section and apply the requested rotation.
            self._center_on_curve(instance)
            if is_col:
                self._rotate_vertical_member(
                    instance,
                    start,
                    getattr(member, "rotation", 0.0),
                )
            else:
                self._set_rotation(instance, getattr(member, "rotation", 0.0))
                if (StructuralFramingUtils is not None
                        and getattr(member, "disallow_end_joins", False)):
                    for end_index in (0, 1):
                        try:
                            StructuralFramingUtils.DisallowJoinAtEnd(
                                instance,
                                end_index,
                            )
                        except Exception:
                            pass

            if getattr(self.config, "track_members", True):
                try:
                    tag_instance(instance, host_info, member)
                except Exception:
                    pass

            placed.append(instance)
            self._last_placed_pairs.append((member, instance))

        return placed

    # ------------------------------------------------------------------
    #  Family resolution
    # ------------------------------------------------------------------

    def _resolve_symbol(self, member):
        """Find the Revit FamilySymbol for a framing member."""
        from Autodesk.Revit.DB import BuiltInCategory

        is_col = getattr(member, "is_column", False)
        primary_cat = (BuiltInCategory.OST_StructuralColumns
                       if is_col
                       else BuiltInCategory.OST_StructuralFraming)
        fallback_cat = (BuiltInCategory.OST_StructuralFraming
                        if is_col
                        else BuiltInCategory.OST_StructuralColumns)

        if member.family_name and member.type_name:
            symbol = find_family_symbol(
                self.doc,
                member.family_name,
                member.type_name,
                primary_cat,
            )
            if symbol is not None:
                return symbol
            symbol = find_family_symbol(
                self.doc,
                member.family_name,
                member.type_name,
                fallback_cat,
            )
            if symbol is not None:
                return symbol

        if self.config.stud_family_name and self.config.stud_type_name:
            symbol = find_family_symbol(
                self.doc,
                self.config.stud_family_name,
                self.config.stud_type_name,
                primary_cat,
            )
            if symbol is not None:
                return symbol
            return find_family_symbol(
                self.doc,
                self.config.stud_family_name,
                self.config.stud_type_name,
                fallback_cat,
            )

        return None

    def get_type_depth(self, family_name, type_name):
        """Read the 'd' dimension from a structural framing family type.

        Returns the depth in feet, or None if unavailable.
        """
        from Autodesk.Revit.DB import BuiltInCategory

        if self.doc is None:
            return None

        if not family_name or not type_name:
            return None

        symbol = find_family_symbol(
            self.doc,
            family_name,
            type_name,
            BuiltInCategory.OST_StructuralFraming,
        )
        if symbol is None:
            return None

        param = symbol.LookupParameter("d")
        if param is not None:
            try:
                val = param.AsDouble()
                if val > 0:
                    return val
            except Exception:
                pass
        return None

    def get_type_width(self, family_name, type_name):
        """Read the 'b' dimension from a structural framing family type.

        Returns the width in feet, or None if unavailable.
        """
        try:
            from Autodesk.Revit.DB import BuiltInCategory
        except Exception:
            return None

        if self.doc is None:
            return None

        if not family_name or not type_name:
            return None

        symbol = find_family_symbol(
            self.doc,
            family_name,
            type_name,
            BuiltInCategory.OST_StructuralFraming,
        )
        if symbol is None:
            return None

        param = symbol.LookupParameter("b")
        if param is not None:
            try:
                val = param.AsDouble()
                if val > 0:
                    return val
            except Exception:
                pass
        return None

    # ------------------------------------------------------------------
    #  Instance configuration
    # ------------------------------------------------------------------

    def _place_column_member(self, member, symbol, level):
        """Place a vertical structural column using point-based placement."""
        from Autodesk.Revit.DB import BuiltInParameter
        from Autodesk.Revit.DB.Structure import StructuralType

        start = member.start_point
        end = member.end_point

        try:
            instance = self.doc.Create.NewFamilyInstance(
                start,
                symbol,
                level,
                StructuralType.Column,
            )
        except Exception:
            return None

        if instance is None:
            return None

        for param_id in (
            BuiltInParameter.FAMILY_TOP_LEVEL_PARAM,
            BuiltInParameter.FAMILY_BASE_LEVEL_PARAM,
        ):
            _set_element_id(instance, param_id, level.Id)

        level_elevation = getattr(level, "Elevation", 0.0)
        base_offset = start.Z - level_elevation
        top_offset = end.Z - level_elevation
        _set_double(
            instance,
            BuiltInParameter.FAMILY_BASE_LEVEL_OFFSET_PARAM,
            base_offset,
        )
        _set_double(
            instance,
            BuiltInParameter.FAMILY_TOP_LEVEL_OFFSET_PARAM,
            top_offset,
        )
        return instance

    @staticmethod
    def _is_vertical_member(member):
        start = getattr(member, "start_point", None)
        end = getattr(member, "end_point", None)
        if start is None or end is None:
            return False
        try:
            dx = abs(end.X - start.X)
            dy = abs(end.Y - start.Y)
            dz = abs(end.Z - start.Z)
        except Exception:
            return False
        return dz > MIN_MEMBER_LENGTH and dx < 1e-6 and dy < 1e-6

    def _center_on_curve(self, instance):
        """Center the cross-section on its analytical curve.

        Sets Y and Z justification to Center so the family's center
        reference planes align with the placement curve.
        """
        from Autodesk.Revit.DB import BuiltInParameter

        # Uniform mode -- both ends share the same justification.
        _set_int(instance, BuiltInParameter.YZ_JUSTIFICATION, 0)

        # Center in both section directions (2 = Center).
        _set_int(instance, BuiltInParameter.Y_JUSTIFICATION, 2)
        _set_int(instance, BuiltInParameter.Z_JUSTIFICATION, 2)

        # Zero all offset values.
        for pid in (
            BuiltInParameter.Y_OFFSET_VALUE,
            BuiltInParameter.Z_OFFSET_VALUE,
            BuiltInParameter.START_Y_OFFSET_VALUE,
            BuiltInParameter.END_Y_OFFSET_VALUE,
            BuiltInParameter.START_Z_OFFSET_VALUE,
            BuiltInParameter.END_Z_OFFSET_VALUE,
        ):
            _set_double(instance, pid, 0.0)

    @staticmethod
    def _set_rotation(instance, angle):
        """Set the absolute cross-section rotation (radians)."""
        from Autodesk.Revit.DB import BuiltInParameter

        _set_double(
            instance,
            BuiltInParameter.STRUCTURAL_BEND_DIR_ANGLE,
            angle,
        )

    def _rotate_vertical_member(self, instance, base_point, angle):
        """Rotate a vertical column around its Z axis."""
        if abs(angle) < 1e-9:
            return

        from Autodesk.Revit.DB import ElementTransformUtils, Line, XYZ

        try:
            axis = Line.CreateBound(base_point, base_point + XYZ.BasisZ)
            ElementTransformUtils.RotateElement(
                self.doc,
                instance.Id,
                axis,
                angle,
            )
        except Exception:
            pass

    @staticmethod
    def _curve_length(start, end):
        """Distance between two XYZ points."""
        try:
            return (end - start).GetLength()
        except Exception:
            return 0.0


# ------------------------------------------------------------------
#  Helpers (module-level, no state)
# ------------------------------------------------------------------

def _set_int(instance, param_id, value):
    """Set an integer parameter, ignoring errors."""
    try:
        p = instance.get_Parameter(param_id)
        if p is not None and not p.IsReadOnly:
            p.Set(value)
    except Exception:
        pass


def _set_double(instance, param_id, value):
    """Set a double parameter, ignoring errors."""
    try:
        p = instance.get_Parameter(param_id)
        if p is not None and not p.IsReadOnly:
            p.Set(value)
    except Exception:
        pass


def _set_element_id(instance, param_id, value):
    """Set an ElementId parameter, ignoring errors."""
    try:
        p = instance.get_Parameter(param_id)
        if p is not None and not p.IsReadOnly:
            p.Set(value)
    except Exception:
        pass