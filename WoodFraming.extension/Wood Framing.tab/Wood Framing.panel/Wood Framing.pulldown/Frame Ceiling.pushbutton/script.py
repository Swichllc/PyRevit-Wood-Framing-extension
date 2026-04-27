# -*- coding: utf-8 -*-
"""Frame Ceiling - Main command script.

Select ceiling hosts and automatically generate ceiling framing
using structural framing families for joists and rim joists.
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

from wf_config import (
    CEILING_DIRECTION_AUTO,
    CEILING_DIRECTION_X,
    CEILING_DIRECTION_Y,
    CEILING_PLACEMENT_ABOVE,
    CEILING_PLACEMENT_CENTER,
    FramingConfig,
    SPACING_16OC,
    SPACING_24OC,
)
from wf_families import get_available_types_flat, parse_family_type_label
from wf_ceiling import CeilingFramingEngine
from wf_tracking import delete_tracked_members_for_hosts

logger = script.get_logger()
output = script.get_output()

_XAML = os.path.join(os.path.dirname(__file__), "FrameCeilingConfig.xaml")
_CFG_PATH = os.path.join(
    os.environ.get("APPDATA", ""),
    "pyRevit",
    "WoodFraming_CeilingLastConfig.json",
)


def _is_ceiling(element):
    try:
        if int(element.Category.Id.IntegerValue) == int(DB.BuiltInCategory.OST_Ceilings):
            return True
    except Exception:
        pass
    if hasattr(DB, 'Ceiling') and isinstance(element, DB.Ceiling):
        return True
    return False


class FrameCeilingDialog(WPFWindow):
    def __init__(self, doc, ceiling_count):
        WPFWindow.__init__(self, _XAML)
        self.doc = doc
        self.result = None
        self._framing_labels = get_available_types_flat(doc)

        self.cb_joist_type.ItemsSource = self._framing_labels
        self.cb_rim_type.ItemsSource = self._framing_labels

        if self._framing_labels:
            self.cb_joist_type.SelectedIndex = 0
            self.cb_rim_type.SelectedIndex = 0

        self.cb_layout_direction.Items.Clear()
        self.cb_layout_direction.Items.Add("Auto (span shorter direction)")
        self.cb_layout_direction.Items.Add("Along local X axis")
        self.cb_layout_direction.Items.Add("Along local Y axis")
        self.cb_layout_direction.SelectedIndex = 0

        self.cb_placement.Items.Clear()
        self.cb_placement.Items.Add("Above ceiling top face")
        self.cb_placement.Items.Add("Center in ceiling layer (legacy)")
        self.cb_placement.SelectedIndex = 0

        self.tb_summary.Text = (
            "Frame {0} ceiling(s). Joists are placed above the ceiling top face by default."
            .format(ceiling_count)
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
            fam, typ = parse_family_type_label(str(joist_sel))
            cfg.stud_family_name = fam
            cfg.stud_type_name = typ

        rim_sel = self.cb_rim_type.SelectedItem
        if rim_sel:
            fam, typ = parse_family_type_label(str(rim_sel))
            cfg.bottom_plate_family_name = fam
            cfg.bottom_plate_type_name = typ

        direction_idx = self.cb_layout_direction.SelectedIndex
        cfg.ceiling_direction_mode = [
            CEILING_DIRECTION_AUTO,
            CEILING_DIRECTION_X,
            CEILING_DIRECTION_Y,
        ][max(0, min(2, direction_idx))]

        placement_idx = self.cb_placement.SelectedIndex
        cfg.ceiling_placement_mode = [
            CEILING_PLACEMENT_ABOVE,
            CEILING_PLACEMENT_CENTER,
        ][max(0, min(1, placement_idx))]

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
                "direction_idx": self.cb_layout_direction.SelectedIndex,
                "placement_idx": self.cb_placement.SelectedIndex,
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

            direction_idx = data.get("direction_idx", 0)
            if direction_idx in (0, 1, 2):
                self.cb_layout_direction.SelectedIndex = direction_idx

            placement_idx = data.get("placement_idx", 0)
            if placement_idx in (0, 1):
                self.cb_placement.SelectedIndex = placement_idx

            if data.get("spacing_24"):
                self.rb_24oc.IsChecked = True
            elif data.get("spacing_custom"):
                self.rb_custom.IsChecked = True
                self.tb_custom_spacing.Text = str(data.get("custom_val", "16"))
            else:
                self.rb_16oc.IsChecked = True
        except Exception:
            pass


def main():
    doc = revit.doc

    available_types = get_available_types_flat(doc)
    if not available_types:
        forms.alert(
            "No Structural Framing families are loaded.\n"
            "Frame Ceiling uses structural framing family types for joists and rim joists.\n"
            "Load a structural framing family before running this command.",
            title="Wood Framing",
        )
        return

    selected = revit.get_selection().elements
    ceilings = [e for e in selected if _is_ceiling(e)]

    if not ceilings:
        try:
            refs = revit.uidoc.Selection.PickObjects(
                ObjectType.Element,
                _CeilingFilter(),
                "Select ceiling hosts to frame",
            )
            ceilings = [doc.GetElement(r.ElementId) for r in refs]
        except Exception:
            return

    if not ceilings:
        forms.alert("No ceilings selected.", title="Wood Framing")
        return

    dlg = FrameCeilingDialog(doc, len(ceilings))
    dlg.ShowDialog()
    cfg = dlg.result
    if cfg is None:
        return

    engine = CeilingFramingEngine(doc, cfg)
    total_placed = 0
    total_ceilings = 0
    deleted_existing = 0

    with revit.Transaction("WF: Frame Ceilings"):
        deleted_existing = delete_tracked_members_for_hosts(
            doc,
            ceilings,
            ("ceiling",),
        )
        for ceiling in ceilings:
            members, ceiling_info = engine.calculate_members(ceiling)
            if ceiling_info is None:
                logger.warning(
                    "Skipped ceiling {0}.".format(ceiling.Id.Value)
                )
                continue
            placed = engine.place_members(members, ceiling_info)
            total_placed += len(placed)
            total_ceilings += 1

    output.print_md(
        "## Ceiling Framing Complete\n"
        "- **Ceilings framed:** {0}\n"
        "- **Previous members replaced:** {1}\n"
        "- **Members placed:** {2}\n"
        "- **Joist direction:** {3}\n"
        "- **Placement:** {4}".format(
            total_ceilings,
            deleted_existing,
            total_placed,
            getattr(cfg, "ceiling_direction_mode", CEILING_DIRECTION_AUTO),
            getattr(cfg, "ceiling_placement_mode", CEILING_PLACEMENT_ABOVE),
        )
    )


class _CeilingFilter(ISelectionFilter):
    def AllowElement(self, element):
        return _is_ceiling(element)

    def AllowReference(self, reference, point):
        return False


if __name__ == "__main__":
    main()
