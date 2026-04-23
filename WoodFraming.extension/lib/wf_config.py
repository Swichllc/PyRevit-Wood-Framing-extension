# -*- coding: utf-8 -*-
"""Framing configuration data model and persistence.

Stores and retrieves framing parameters such as stud spacing,
family types, plate options, and per-project overrides.
"""

import json
import os


# Default stud spacings in INCHES (converted to feet at usage)
SPACING_16OC = 16.0
SPACING_24OC = 24.0

# Plate options
SINGLE_TOP_PLATE = 1
DOUBLE_TOP_PLATE = 2

# Host layer selection
LAYER_MODE_CORE_CENTER = "core_center"
LAYER_MODE_STRUCTURAL = "structural"
LAYER_MODE_THICKEST = "thickest"

# Wall base elevation source
WALL_BASE_MODE_WALL = "wall_base"
WALL_BASE_MODE_SUPPORT_TOP = "support_top"

# Ceiling framing controls
CEILING_DIRECTION_AUTO = "auto"
CEILING_DIRECTION_X = "x_axis"
CEILING_DIRECTION_Y = "y_axis"
CEILING_PLACEMENT_ABOVE = "above_top_face"
CEILING_PLACEMENT_CENTER = "center_in_layer"

# Default member actual dimensions in INCHES
LUMBER_ACTUAL = {
    "2x2": (1.5, 1.5),
    "2x3": (1.5, 2.5),
    "2x4": (1.5, 3.5),
    "2x6": (1.5, 5.5),
    "2x8": (1.5, 7.25),
    "2x10": (1.5, 9.25),
    "2x12": (1.5, 11.25),
}


class FramingConfig(object):
    """Holds all framing configuration for a wall/floor/roof/ceiling operation."""

    def __init__(self):
        # Stud settings
        self.stud_spacing = SPACING_16OC         # inches
        self.stud_family_name = None              # str - Revit family name
        self.stud_type_name = None                # str - Revit type name

        # Bottom plate
        self.bottom_plate_family_name = None
        self.bottom_plate_type_name = None
        self.bottom_plate_count = 1

        # Top plate
        self.top_plate_family_name = None
        self.top_plate_type_name = None
        self.top_plate_count = DOUBLE_TOP_PLATE   # default double

        # Header
        self.header_family_name = None
        self.header_type_name = None
        self.header_count = 2                     # double header (standard)

        # Mid plates
        self.include_mid_plates = True
        self.mid_plate_interval_ft = 8.0

        # Sill plate (below windows)
        self.sill_plate_family_name = None
        self.sill_plate_type_name = None

        # General options
        self.include_corner_studs = True
        self.include_cripple_studs = True
        self.include_king_studs = True
        self.include_jack_studs = True

        # Offsets (feet)
        self.wall_center_offset = 0.0

        # Host layer rules
        self.wall_layer_mode = LAYER_MODE_STRUCTURAL
        self.floor_layer_mode = LAYER_MODE_STRUCTURAL
        self.ceiling_layer_mode = LAYER_MODE_STRUCTURAL
        self.ceiling_direction_mode = CEILING_DIRECTION_AUTO
        self.ceiling_placement_mode = CEILING_PLACEMENT_ABOVE
        self.roof_layer_mode = LAYER_MODE_STRUCTURAL

        # Wall base reference source
        self.wall_base_mode = WALL_BASE_MODE_WALL
        self.wall_base_override_z = None
        self.wall_base_support_element_id = None

        # Roof framing options
        self.include_collar_ties = True
        self.include_ceiling_joists = True
        self.include_roof_kickers = True
        self.roof_edge_family_name = None
        self.roof_edge_type_name = None

        # Generated-member ownership tracking
        self.track_members = True

    @property
    def stud_spacing_ft(self):
        """Stud spacing converted to feet (Revit internal units)."""
        return self.stud_spacing / 12.0

    def to_dict(self):
        """Serialize config to dictionary for JSON persistence."""
        return {
            "stud_spacing": self.stud_spacing,
            "stud_family_name": self.stud_family_name,
            "stud_type_name": self.stud_type_name,
            "bottom_plate_family_name": self.bottom_plate_family_name,
            "bottom_plate_type_name": self.bottom_plate_type_name,
            "bottom_plate_count": self.bottom_plate_count,
            "top_plate_family_name": self.top_plate_family_name,
            "top_plate_type_name": self.top_plate_type_name,
            "top_plate_count": self.top_plate_count,
            "header_family_name": self.header_family_name,
            "header_type_name": self.header_type_name,
            "header_count": self.header_count,
            "include_mid_plates": self.include_mid_plates,
            "mid_plate_interval_ft": self.mid_plate_interval_ft,
            "sill_plate_family_name": self.sill_plate_family_name,
            "sill_plate_type_name": self.sill_plate_type_name,
            "include_corner_studs": self.include_corner_studs,
            "include_cripple_studs": self.include_cripple_studs,
            "include_king_studs": self.include_king_studs,
            "include_jack_studs": self.include_jack_studs,
            "wall_center_offset": self.wall_center_offset,
            "wall_layer_mode": self.wall_layer_mode,
            "floor_layer_mode": self.floor_layer_mode,
            "ceiling_layer_mode": self.ceiling_layer_mode,
            "ceiling_direction_mode": self.ceiling_direction_mode,
            "ceiling_placement_mode": self.ceiling_placement_mode,
            "roof_layer_mode": self.roof_layer_mode,
            "wall_base_mode": self.wall_base_mode,
            "wall_base_override_z": self.wall_base_override_z,
            "wall_base_support_element_id": self.wall_base_support_element_id,
            "include_collar_ties": self.include_collar_ties,
            "include_ceiling_joists": self.include_ceiling_joists,
            "include_roof_kickers": self.include_roof_kickers,
            "roof_edge_family_name": self.roof_edge_family_name,
            "roof_edge_type_name": self.roof_edge_type_name,
            "track_members": self.track_members,
        }

    @classmethod
    def from_dict(cls, data):
        """Deserialize config from dictionary."""
        cfg = cls()
        for key, value in data.items():
            if hasattr(cfg, key):
                setattr(cfg, key, value)
        return cfg

    def save(self, filepath):
        """Persist config to a JSON file."""
        with open(filepath, "w") as f:
            json.dump(self.to_dict(), f, indent=2)

    @classmethod
    def load(cls, filepath):
        """Load config from a JSON file. Returns default if file missing."""
        if not os.path.exists(filepath):
            return cls()
        with open(filepath, "r") as f:
            data = json.load(f)
        return cls.from_dict(data)
