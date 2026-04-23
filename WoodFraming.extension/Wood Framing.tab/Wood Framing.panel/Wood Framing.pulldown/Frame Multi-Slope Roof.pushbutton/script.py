# -*- coding: utf-8 -*-
"""Frame multi-slope roofs with the clean V2 engine.

This command uses the new per-field planner in `wf_roof_v2.py` and places real
framing members without calling the legacy multi-slope logic in `wf_roof.py`.
"""

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
from wf_roof_v2 import RoofFramingEngineV2, _placement_fields_for_plan, V2_BUILD_TAG


output = script.get_output()
_XAML = script.get_bundle_file("FrameMultiSlopeV2Config.xaml")
_CFG_PATH = os.path.join(
    os.environ.get("APPDATA", ""),
    "pyRevit", "WoodFraming_RoofLastConfig.json",
)


class _RoofFilter(ISelectionFilter):
    def AllowElement(self, elem):
        return isinstance(elem, DB.RoofBase)

    def AllowReference(self, ref, pt):
        return False


class FrameMultiSlopeV2Dialog(WPFWindow):
    def __init__(self, doc):
        WPFWindow.__init__(self, _XAML)
        self.doc = doc
        self.result = None

        self._framing_labels = get_available_types_flat(doc)
        self.cb_rafter_type.ItemsSource = self._framing_labels
        self.cb_ridge_type.ItemsSource = self._framing_labels
        self.cb_edge_type.ItemsSource = self._framing_labels

        if self._framing_labels:
            self.cb_rafter_type.SelectedIndex = 0
            self.cb_ridge_type.SelectedIndex = 0
            self.cb_edge_type.SelectedIndex = 0

        self.btn_ok.Click += self._on_ok
        self.btn_cancel.Click += self._on_cancel
        self.rb_custom_sp.Checked += self._on_custom_checked
        self.rb_custom_sp.Unchecked += self._on_custom_unchecked

        self._restore_last()

    def _on_custom_checked(self, sender, args):
        self.tb_custom_spacing.IsEnabled = True

    def _on_custom_unchecked(self, sender, args):
        self.tb_custom_spacing.IsEnabled = False

    def _on_ok(self, sender, args):
        config = FramingConfig()

        rafter_sel = self.cb_rafter_type.SelectedItem
        if rafter_sel:
            family_name, type_name = parse_family_type_label(str(rafter_sel))
            config.stud_family_name = family_name
            config.stud_type_name = type_name

        ridge_sel = self.cb_ridge_type.SelectedItem
        if ridge_sel:
            family_name, type_name = parse_family_type_label(str(ridge_sel))
            config.header_family_name = family_name
            config.header_type_name = type_name

        edge_sel = self.cb_edge_type.SelectedItem
        if edge_sel:
            family_name, type_name = parse_family_type_label(str(edge_sel))
            config.roof_edge_family_name = family_name
            config.roof_edge_type_name = type_name

        if self.rb_16oc.IsChecked:
            config.stud_spacing = SPACING_16OC
        elif self.rb_24oc.IsChecked:
            config.stud_spacing = SPACING_24OC
        else:
            try:
                config.stud_spacing = float(self.tb_custom_spacing.Text)
            except Exception:
                config.stud_spacing = SPACING_16OC

        config.include_collar_ties = False
        config.include_ceiling_joists = False
        config.include_roof_kickers = False

        self.result = {
            "config": config,
            "mode": "stick" if self.rb_stick.IsChecked else "truss",
        }
        self._save_last(config, self.result["mode"])
        self.Close()

    def _on_cancel(self, sender, args):
        self.result = None
        self.Close()

    def _save_last(self, config, mode):
        try:
            data = config.to_dict()
            data["_roof_mode"] = mode
            data["_rafter_label"] = str(self.cb_rafter_type.SelectedItem or "")
            data["_ridge_label"] = str(self.cb_ridge_type.SelectedItem or "")
            data["_edge_label"] = str(self.cb_edge_type.SelectedItem or "")
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

            mode = data.get("_roof_mode", "stick")
            if mode == "truss":
                self.rb_truss.IsChecked = True
            else:
                self.rb_stick.IsChecked = True

            rafter_label = data.get("_rafter_label", "")
            if rafter_label and rafter_label in self._framing_labels:
                self.cb_rafter_type.SelectedItem = rafter_label

            ridge_label = data.get("_ridge_label", "")
            if ridge_label and ridge_label in self._framing_labels:
                self.cb_ridge_type.SelectedItem = ridge_label

            edge_label = data.get("_edge_label", "")
            if edge_label and edge_label in self._framing_labels:
                self.cb_edge_type.SelectedItem = edge_label
            elif ridge_label and ridge_label in self._framing_labels:
                self.cb_edge_type.SelectedItem = ridge_label

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


def _roof_id(roof):
    roof_id = getattr(roof, "Id", None)
    if roof_id is None:
        return "?"
    return getattr(roof_id, "Value", getattr(roof_id, "IntegerValue", "?"))


def _select_roofs(doc):
    selected = revit.get_selection().elements
    roofs = [element for element in selected if isinstance(element, DB.RoofBase)]
    if roofs:
        return roofs

    try:
        refs = revit.uidoc.Selection.PickObjects(
            ObjectType.Element,
            _RoofFilter(),
            "Select roofs to frame with the multi-slope engine",
        )
    except Exception:
        return []
    return [doc.GetElement(ref.ElementId) for ref in refs]


def _print_plan_warnings(roof_id, plan):
    if plan is None or not getattr(plan, "warnings", None):
        return
    for warning in plan.warnings:
        output.print_md("- Roof {0} warning: {1}".format(roof_id, warning))


def main():
    doc = revit.doc
    if not get_available_types_flat(doc):
        forms.alert(
            "No structural framing families are loaded. Load a framing family before running this command.",
            title="Frame Multi-Slope Roof",
        )
        return

    dialog = FrameMultiSlopeV2Dialog(doc)
    dialog.ShowDialog()
    if dialog.result is None:
        return
    if dialog.result["mode"] == "truss":
        forms.alert(
            "Frame Multi-Slope Roof currently supports stick-style framing only.",
            title="Frame Multi-Slope Roof",
        )
        return
    config = dialog.result["config"]

    roofs = _select_roofs(doc)
    if not roofs:
        forms.alert("No roofs selected.", title="Frame Multi-Slope Roof")
        return

    engine = RoofFramingEngineV2(doc, config)
    output.print_md(
        "## Frame Multi-Slope Roof\n"
        "- **Rafter type:** {0} : {1}\n"
        "- **Ridge board type:** {2} : {3}\n"
        "- **Eave / border type:** {4} : {5}\n"
        "- **Spacing:** {6:.2f} in OC\n"
        "- **Current scope:** rafters, ridge boards, and border members\n"
        "- **Engine build:** {7}".format(
            config.stud_family_name,
            config.stud_type_name,
            config.header_family_name or config.stud_family_name,
            config.header_type_name or config.stud_type_name,
            config.roof_edge_family_name or config.header_family_name or config.stud_family_name,
            config.roof_edge_type_name or config.header_type_name or config.stud_type_name,
            config.stud_spacing,
            V2_BUILD_TAG,
        )
    )
    total_roofs = 0
    total_members = 0
    total_placed = 0
    skipped = 0
    errors = []

    with revit.Transaction("WF: Frame Multi-Slope Roof"):
        for roof in roofs:
            roof_id = _roof_id(roof)
            try:
                members, roof_info = engine.calculate_members(roof)
            except Exception as exc:
                errors.append("Roof {0} calc error: {1}".format(roof_id, exc))
                continue

            if roof_info is None:
                errors.append("Roof {0}: analyze_roof_host returned None".format(roof_id))
                continue

            plan = getattr(roof_info, "v2_plan", None)
            if plan is None or not plan.supported:
                skipped += 1
                reason = "No supported multi-slope framing plan was created."
                if plan is not None and getattr(plan, "warnings", None):
                    reason = plan.warnings[0]
                errors.append("Roof {0} skipped: {1}".format(roof_id, reason))
                continue

            total_members += len(members)
            if not members:
                skipped += 1
                errors.append("Roof {0} planned but no members were generated.".format(roof_id))
                _print_plan_warnings(roof_id, plan)
                continue

            try:
                placed = engine.place_members(members, roof_info)
            except Exception as exc:
                errors.append("Roof {0} place error: {1}".format(roof_id, exc))
                placed = []

            total_roofs += 1
            total_placed += len(placed)
            analyzed_field_count = len(getattr(plan, "fields", []) or [])
            try:
                placement_field_count = len(_placement_fields_for_plan(plan))
            except Exception:
                placement_field_count = 0
            output.print_md(
                "- Roof {0}: {1} members generated, {2} placed, {3} placement fields selected ({4} analyzed)".format(
                    roof_id,
                    len(members),
                    len(placed),
                    placement_field_count,
                    analyzed_field_count,
                )
            )
            type_counts = {}
            for member in members:
                member_type = getattr(member, "member_type", "UNKNOWN")
                type_counts[member_type] = type_counts.get(member_type, 0) + 1
            if type_counts:
                parts = []
                for member_type, count in sorted(type_counts.items()):
                    parts.append("{0}={1}".format(member_type, count))
                output.print_md("  - Types: " + ", ".join(parts))
            _print_plan_warnings(roof_id, plan)

    output.print_md(
        "## Frame Multi-Slope Roof Complete\n"
        "- **Roofs framed:** {0}\n"
        "- **Roofs skipped:** {1}\n"
        "- **Members generated:** {2}\n"
        "- **Members placed:** {3}".format(
            total_roofs,
            skipped,
            total_members,
            total_placed,
        )
    )

    if errors:
        output.print_md("\n### Notes")
        for line in errors:
            output.print_md("- " + str(line))


if __name__ == "__main__":
    main()