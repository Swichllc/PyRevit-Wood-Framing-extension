# -*- coding: utf-8 -*-
"""Single-slope roof framing for shed roofs."""

import os
import sys
import json

_ext_dir = os.path.dirname(os.path.dirname(os.path.dirname(
    os.path.dirname(__file__)
)))
while _ext_dir and not _ext_dir.lower().endswith(".extension"):
    _parent = os.path.dirname(_ext_dir)
    if _parent == _ext_dir:
        break
    _ext_dir = _parent
_lib_dir = os.path.join(_ext_dir, "lib")
if _lib_dir not in sys.path:
    sys.path.insert(0, _lib_dir)

from pyrevit import revit, DB, script, forms
from pyrevit.forms import WPFWindow
from Autodesk.Revit.UI.Selection import ObjectType, ISelectionFilter

from wf_config import FramingConfig, SPACING_16OC, SPACING_24OC
from wf_families import get_available_types_flat, parse_family_type_label
from wf_roof import RoofFramingEngine
from wf_tracking import delete_tracked_members_for_hosts


output = script.get_output()
_XAML = os.path.join(os.path.dirname(__file__), "FrameRoofConfig.xaml")
_CFG_PATH = os.path.join(
    os.environ.get("APPDATA", ""),
    "pyRevit", "WoodFraming_RoofLastConfig.json",
)


class _RoofFilter(ISelectionFilter):
    def AllowElement(self, elem):
        return isinstance(elem, DB.RoofBase)

    def AllowReference(self, ref, pt):
        return False


class FrameRoofDialog(WPFWindow):
    def __init__(self, doc):
        WPFWindow.__init__(self, _XAML)
        self.doc = doc
        self.result = None

        self._framing_labels = get_available_types_flat(doc)
        self.cb_rafter_type.ItemsSource = self._framing_labels
        self.cb_ridge_type.ItemsSource = self._framing_labels

        if self._framing_labels:
            self.cb_rafter_type.SelectedIndex = 0
            self.cb_ridge_type.SelectedIndex = 0

        self.rb_custom_sp.Checked += self._on_custom_checked
        self.rb_custom_sp.Unchecked += self._on_custom_unchecked
        self.btn_ok.Click += self._on_ok
        self.btn_cancel.Click += self._on_cancel

        self._restore_last()

    def _on_custom_checked(self, sender, args):
        self.tb_custom_spacing.IsEnabled = True

    def _on_custom_unchecked(self, sender, args):
        self.tb_custom_spacing.IsEnabled = False

    def _on_ok(self, sender, args):
        cfg = FramingConfig()

        rafter_sel = self.cb_rafter_type.SelectedItem
        if rafter_sel:
            family_name, type_name = parse_family_type_label(str(rafter_sel))
            cfg.stud_family_name = family_name
            cfg.stud_type_name = type_name

        ridge_sel = self.cb_ridge_type.SelectedItem
        if ridge_sel:
            family_name, type_name = parse_family_type_label(str(ridge_sel))
            cfg.header_family_name = family_name
            cfg.header_type_name = type_name

        if self.rb_16oc.IsChecked:
            cfg.stud_spacing = SPACING_16OC
        elif self.rb_24oc.IsChecked:
            cfg.stud_spacing = SPACING_24OC
        else:
            try:
                cfg.stud_spacing = float(self.tb_custom_spacing.Text)
            except Exception:
                cfg.stud_spacing = SPACING_16OC

        cfg.include_collar_ties = False
        cfg.include_ceiling_joists = False
        cfg.include_roof_kickers = False

        self.result = {"config": cfg}
        self._save_last(cfg)
        self.Close()

    def _on_cancel(self, sender, args):
        self.result = None
        self.Close()

    def _save_last(self, cfg):
        try:
            data = cfg.to_dict()
            data["_rafter_label"] = str(self.cb_rafter_type.SelectedItem or "")
            data["_ridge_label"] = str(self.cb_ridge_type.SelectedItem or "")
            directory = os.path.dirname(_CFG_PATH)
            if not os.path.exists(directory):
                os.makedirs(directory)
            with open(_CFG_PATH, "w") as stream:
                json.dump(data, stream, indent=2)
        except Exception:
            pass

    def _restore_last(self):
        try:
            if not os.path.exists(_CFG_PATH):
                return
            with open(_CFG_PATH, "r") as stream:
                data = json.load(stream)

            rafter_label = data.get("_rafter_label", "")
            if rafter_label and rafter_label in self._framing_labels:
                self.cb_rafter_type.SelectedItem = rafter_label

            ridge_label = data.get("_ridge_label", "")
            if ridge_label and ridge_label in self._framing_labels:
                self.cb_ridge_type.SelectedItem = ridge_label

            spacing = data.get("stud_spacing", SPACING_16OC)
            if spacing == SPACING_16OC:
                self.rb_16oc.IsChecked = True
            elif spacing == SPACING_24OC:
                self.rb_24oc.IsChecked = True
            else:
                self.rb_custom_sp.IsChecked = True
                self.tb_custom_spacing.Text = str(spacing)
        except Exception:
            pass


def main():
    doc = revit.doc

    framing_types = get_available_types_flat(doc)
    if not framing_types:
        forms.alert(
            "No structural framing families are loaded.\n"
            "Load a framing family before running this command.",
            title="Wood Framing",
        )
        return

    selected = revit.get_selection().elements
    roofs = [element for element in selected if isinstance(element, DB.RoofBase)]

    if not roofs:
        try:
            refs = revit.uidoc.Selection.PickObjects(
                ObjectType.Element,
                _RoofFilter(),
                "Select single-slope roofs to frame",
            )
            roofs = [doc.GetElement(ref.ElementId) for ref in refs]
        except Exception:
            return

    if not roofs:
        forms.alert("No roofs selected.", title="Wood Framing")
        return

    dialog = FrameRoofDialog(doc)
    dialog.ShowDialog()
    if dialog.result is None:
        return

    config = dialog.result["config"]
    engine = RoofFramingEngine(doc, config)

    total_placed = 0
    total_calculated = 0
    total_roofs = 0
    skipped_roofs = 0
    deleted_existing = 0
    errors = []

    with revit.Transaction("WF: Frame Single-Slope Roofs"):
        deleted_existing = delete_tracked_members_for_hosts(
            doc,
            roofs,
            ("roof",),
        )
        for roof in roofs:
            try:
                members, roof_info = engine.calculate_members(roof, mode="stick")
            except Exception as calc_err:
                errors.append(
                    "Roof {0} calc error: {1}".format(roof.Id.Value, calc_err)
                )
                continue

            if roof_info is None:
                errors.append(
                    "Roof {0}: analyze_roof_host returned None".format(roof.Id.Value)
                )
                continue

            if not getattr(roof_info, "single_slope_supported", True):
                errors.append(
                    "Roof {0} skipped: {1}".format(
                        roof.Id.Value,
                        getattr(
                            roof_info,
                            "single_slope_support_reason",
                            "Single-slope roof framing currently supports shed roofs only.",
                        ),
                    )
                )
                skipped_roofs += 1
                continue

            total_calculated += len(members)

            for plane_index, plane in enumerate(roof_info.planes):
                output.print_md(
                    "  - Plane {0}: normal=({1:.3f},{2:.3f},{3:.3f}), "
                    "bounds=({4:.1f},{5:.1f},{6:.1f},{7:.1f}), loops={8}".format(
                        plane_index,
                        plane.normal.X, plane.normal.Y, plane.normal.Z,
                        plane.bounds[0], plane.bounds[1],
                        plane.bounds[2], plane.bounds[3],
                        len(plane.boundary_loops_local),
                    )
                )

            try:
                placed = engine.place_members(members, roof_info)
            except Exception as place_err:
                errors.append(
                    "Roof {0} place error: {1}".format(roof.Id.Value, place_err)
                )
                placed = []

            total_placed += len(placed)
            total_roofs += 1

    output.print_md(
        "## Single-Slope Roof Framing Complete\n"
        "- **Roofs framed:** {0}\n"
        "- **Roofs skipped:** {1}\n"
        "- **Previous members replaced:** {2}\n"
        "- **Members calculated:** {3}\n"
        "- **Members placed:** {4}".format(
            total_roofs,
            skipped_roofs,
            deleted_existing,
            total_calculated,
            total_placed,
        )
    )

    if errors:
        output.print_md("\n### Errors")
        for line in errors:
            output.print_md("- " + str(line))

    if total_calculated > 0 and total_placed == 0:
        output.print_md(
            "\n> **Warning:** Members were calculated but none could be "
            "placed. Check that the selected family types exist in the "
            "project and that the roof has a valid level assignment."
        )
    elif total_calculated == 0 and total_roofs > 0:
        output.print_md(
            "\n> **Warning:** No framing members were generated. "
            "Check the plane diagnostics above for geometry details."
        )
    elif skipped_roofs and total_roofs == 0:
        output.print_md(
            "\n> **Note:** Multi-slope roofs are intentionally blocked while "
            "the roof framing workflow stays focused on the single-slope tool."
        )


if __name__ == "__main__":
    main()
