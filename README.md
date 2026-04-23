# swich-wood-framing

swich-wood-framing is a pyRevit extension for Autodesk Revit that automates wood framing workflows for walls, floors, ceilings, and roofs.

The Git repository is named `swich-wood-framing`. The extension loaded inside pyRevit keeps its current package identity: `WoodFraming.extension`, with the `Wood Framing` tab visible in Revit.

## Included tools

- Frame Wall
- Frame Floor
- Frame Ceiling
- Single-Slope Roof Framing
- Plan Multi-Slope Roof Framing
- Frame Multi-Slope Roof V2
- Multi Frame
- Update Frame
- Delete Frame
- Split Sheathing
- Generate BOM
- Number Members
- Load Families
- WF Settings

## Repository layout

```text
WoodFraming.extension/
  extension.json
  lib/
  Wood Framing.tab/
```

- `extension.json` defines the pyRevit extension metadata.
- `lib/` contains the shared framing, geometry, family, floor, ceiling, roof, and configuration logic.
- `Wood Framing.tab/` contains the pyRevit buttons grouped by Documentation, Framing, Modify, Openings, and Settings panels.

## Prerequisites

- Autodesk Revit
- pyRevit
- Access to the Revit families and project templates used by your office workflow

## Installation

1. Make sure pyRevit is installed and working in Revit.
2. Place this repository inside a folder that pyRevit uses as an extensions root.
3. Keep `WoodFraming.extension` directly inside that extensions root.
4. Reload pyRevit or restart Revit.

## Development notes

- This repository keeps the current extension branding unchanged on purpose.
- The root markdown document that lives beside the extension in this workspace is intentionally excluded from Git.
- Offline roof regression checks are available with `python run_roof_regressions.py`.
- The current offline harness covers shed-roof profile classification, scanline clipping, and rake-board inside-side selection without launching Revit.

## Next Git step

After creating the remote repository, add it as `origin` and push the local `main` branch.