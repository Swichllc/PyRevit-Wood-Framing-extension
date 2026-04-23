# -*- coding: utf-8 -*-
"""Frame Floor - Main command script."""

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
from wf_floor import FloorFramingEngine

logger = script.get_logger()
output = script.get_output()

_XAML = os.path.join(os.path.dirname(__file__), "FrameFloorConfig.xaml")
_CFG_PATH = os.path.join(
    os.environ.get("APPDATA", ""),
    "pyRevit",
    "WoodFraming_FloorLastConfig.json",
)


class FrameFloorDialog(WPFWindow):
    def __init__(self, doc, floor_count):
        WPFWindow.__init__(self, _XAML)
        self.doc = doc
        self.result = None
        self._framing_labels = get_available_types_flat(doc)

        self.cb_joist_type.ItemsSource = self._framing_labels
        self.cb_rim_type.ItemsSource = self._framing_labels

        if self._framing_labels:
            self.cb_joist_type.SelectedIndex = 0
            self.cb_rim_type.SelectedIndex = 0

        self.tb_summary.Text = (
            "Frame {0} floor(s). Joists and rim joists use structural framing families."
            .format(floor_count)
        )

        self.rb_custom.Checked += self._on_custom_checked
        self.rb_custom.Unchecked += self._on_custom_unchecked
        self.btn_ok.Click += self._on_ok
        self.btn_cancel.Click += self._on_cancel

        self._restore_last()

    def _on_custom_checked(self, sender, args):
        self.tb_custom_spacing.IsEnabled = True

    def _on_custom_unchecked(self, sender, args):
        self.tb_custom_spacing.IsEnabled = False

    def _on_ok(self, sender, args):
        self.result = self._build_config()
        self._save_last()
        self.Close()

    def _on_cancel(self, sender, args):
        self.result = None
        self.Close()

    def _build_config(self):
        cfg = FramingConfig()

        if self.rb_24oc.IsChecked:
            cfg.stud_spacing = SPACING_24OC
        elif self.rb_custom.IsChecked:
            try:
                cfg.stud_spacing = float(self.tb_custom_spacing.Text)
            except Exception:
                cfg.stud_spacing = SPACING_16OC
        else:
            cfg.stud_spacing = SPACING_16OC

        joist_sel = self.cb_joist_type.SelectedItem
        if joist_sel:
            family_name, type_name = parse_family_type_label(str(joist_sel))
            cfg.stud_family_name = family_name
            cfg.stud_type_name = type_name

        rim_sel = self.cb_rim_type.SelectedItem
        if rim_sel:
            family_name, type_name = parse_family_type_label(str(rim_sel))
            cfg.bottom_plate_family_name = family_name
            cfg.bottom_plate_type_name = type_name

        return cfg

    def _save_last(self):
        try:
            data = {
                "joist_label": str(self.cb_joist_type.SelectedItem or ""),
                "rim_label": str(self.cb_rim_type.SelectedItem or ""),
                "spacing_16": bool(self.rb_16oc.IsChecked),
                "spacing_24": bool(self.rb_24oc.IsChecked),
                "spacing_custom": bool(self.rb_custom.IsChecked),
                "custom_val": self.tb_custom_spacing.Text,
            }
            cfg_dir = os.path.dirname(_CFG_PATH)
            if not os.path.exists(cfg_dir):
                os.makedirs(cfg_dir)
            with open(_CFG_PATH, "w") as cfg_file:
                json.dump(data, cfg_file)
        except Exception:
            pass

    def _restore_last(self):
        try:
            if not os.path.exists(_CFG_PATH):
                return
            with open(_CFG_PATH, "r") as cfg_file:
                data = json.load(cfg_file)

            joist_label = data.get("joist_label", "")
            if joist_label in self._framing_labels:
                self.cb_joist_type.SelectedItem = joist_label

            rim_label = data.get("rim_label", "")
            if rim_label in self._framing_labels:
                self.cb_rim_type.SelectedItem = rim_label

            if data.get("spacing_24"):
                self.rb_24oc.IsChecked = True
            elif data.get("spacing_custom"):
                self.rb_custom.IsChecked = True
                self.tb_custom_spacing.Text = str(data.get("custom_val", "16"))
            else:
                self.rb_16oc.IsChecked = True
        except Exception:
            pass


class _FloorFilter(ISelectionFilter):
    def AllowElement(self, element):
        return isinstance(element, DB.Floor)

    def AllowReference(self, reference, point):
        return False


def main():
    doc = revit.doc

    available_types = get_available_types_flat(doc)
    if not available_types:
        forms.alert(
            "No structural framing families are loaded.\n"
            "Load a framing family before running this command.",
            title="Wood Framing",
        )
        return

    selected = revit.get_selection().elements
    floors = [element for element in selected if isinstance(element, DB.Floor)]

    if not floors:
        try:
            refs = revit.uidoc.Selection.PickObjects(
                ObjectType.Element,
                _FloorFilter(),
                "Select floors to frame",
            )
            floors = [doc.GetElement(ref.ElementId) for ref in refs]
        except Exception:
            return

    if not floors:
        forms.alert("No floors selected.", title="Wood Framing")
        return

    dialog = FrameFloorDialog(doc, len(floors))
    dialog.ShowDialog()
    config = dialog.result
    if config is None:
        return

    engine = FloorFramingEngine(doc, config)
    total_placed = 0
    total_floors = 0

    with revit.Transaction("WF: Frame Floors"):
        for floor in floors:
            members, floor_info = engine.calculate_members(floor)
            if floor_info is None:
                logger.warning("Skipped floor {0}.".format(floor.Id.Value))
                continue
            placed = engine.place_members(members, floor_info)
            total_placed += len(placed)
            total_floors += 1

    output.print_md(
        "## Floor Framing Complete\n"
        "- **Floors framed:** {0}\n"
        "- **Members placed:** {1}".format(total_floors, total_placed)
    )


if __name__ == "__main__":
    main()
