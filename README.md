# swich-wood-framing

`swich-wood-framing` is a pyRevit extension for Autodesk Revit wood framing workflows.

This repository contains one pyRevit extension folder:

```text
WoodFraming.extension/
```

When loaded in Revit, the extension adds a `Wood Framing` tab with a `Wood Framing` panel and `Wood Framing` pulldown.

## Current status

- Work in progress.
- Tested only in Autodesk Revit 2026.
- Not verified in any other Revit version.
- Not perfect. Review the model output before using it for production, estimating, fabrication, coordination, or construction decisions.

## Disclaimer

Use this repository and pyRevit extension at your own risk. The author and contributors accept no responsibility for model changes, data loss, incorrect quantities, incorrect framing, project delays, construction issues, or any other direct or indirect damages from using this work.

## Active tools

These are the active pyRevit button titles currently defined by the enabled `*.pushbutton` bundles in this repo:

| Tool name in Revit | Bundle folder |
| --- | --- |
| `Wall Framing` | `Wall Framing.pushbutton` |
| `Floor Framing` | `Frame Floor.pushbutton` |
| `Ceiling Framing` | `Frame Ceiling.pushbutton` |
| `Single-Slope Roof Framing` | `Frame Roof.pushbutton` |
| `Multi-Slope Roof` | `Frame Multi-Slope Roof.pushbutton` |
| `Split Sheathing` | `Split Sheathing.pushbutton` |
| `Number Members` | `Number Members.pushbutton` |
| `Material List` | `Generate BOM.pushbutton` |

Disabled or legacy folders may exist in the repo, but they are not listed above as active tools.

## Prerequisites

- Autodesk Revit 2026.
- pyRevit installed and attached to Revit 2026.
- Revit family/content setup that matches the framing workflow used by the tools.

## Installation option 1: install into the pyRevit extensions folder

Use this option when you want pyRevit to load the extension from its default user extensions location.

1. Close Revit.
2. Open this folder in Windows Explorer:

   ```text
   %APPDATA%\pyRevit\Extensions
   ```

   Create the `Extensions` folder if it does not exist.

3. Copy the `WoodFraming.extension` folder from this repo into that folder.
4. Confirm the final structure looks like this:

   ```text
   %APPDATA%\pyRevit\Extensions\
     WoodFraming.extension\
       extension.json
       lib\
       Wood Framing.tab\
   ```

5. Open Revit 2026.
6. Reload pyRevit or restart Revit if the `Wood Framing` tab does not appear.

## Installation option 2: connect this repo through pyRevit settings

Use this option when you want to keep the repo in place and have pyRevit load it from this working folder.

1. Keep the repo folder in a stable location.
2. In Revit, open pyRevit settings.
3. Add the parent folder that contains `WoodFraming.extension` to the custom extension directories list.

   For this checkout, add:

   ```text
   d:\Temp projects\009 - wood framing
   ```

   Do not add the `WoodFraming.extension` folder itself. Add the folder that contains it.

4. Save the settings.
5. Reload pyRevit or restart Revit.
6. Look for the `Wood Framing` tab and `Wood Framing` pulldown.

## Repository layout

```text
WoodFraming.extension/
  extension.json
  lib/
  Wood Framing.tab/
```

- `extension.json` contains the pyRevit extension metadata.
- `lib/` contains shared Python modules for framing, geometry, family handling, schedules, tracking, host analysis, floors, ceilings, roofs, and wall framing.
- `Wood Framing.tab/` contains the pyRevit ribbon structure and active tool buttons.

## pyRevit references checked

- pyRevit official API docs define the default third-party extension location as `%APPDATA%\pyRevit\Extensions`: <https://docs.pyrevitlabs.io/reference/pyrevit/>
- pyRevit official user configuration docs describe user extension root directories: <https://docs.pyrevitlabs.io/reference/pyrevit/userconfig/>
- pyRevit extension docs identify `.extension`, `.tab`, `.panel`, `.pulldown`, and `.pushbutton` bundle naming: <https://docs.pyrevitlabs.io/reference/pyrevit/extensions/>
- pyRevit manual extension guidance says to add the directory containing the extension folder: <https://pyrevitlabs.notion.site/Install-Extensions-0753ab78c0ce46149f962acc50892491>
