# FeatureWorks STEP Importer (VBA macro)

A SOLIDWORKS VBA macro that adds a toolbar button: click it, pick a STEP
file (it'll keep asking for more until you say no), and it runs FeatureWorks
feature recognition on each one with:

- **Recognition mode:** Automatic (`RecognizeFeatureAutomatic`)
- **Feature type:** Standard features only (not sheet metal)
- **Auto-recognized types:** Extrude/boss/cut, Revolve, Holes, Fillets &
  Chamfers, Ribs — i.e. everything available in automatic mode **except
  Volume**. Draft isn't included either, but that's automatic for free:
  FeatureWorks has no automatic-mode option for Draft at all — it's only
  recognized interactively.

Each recognized part is saved as a native `.SLDPRT` next to the source STEP
file (same name, `.SLDPRT` extension), then closed.

## Prerequisites

- FeatureWorks must be licensed and enabled: **Tools > Add-ins**, tick
  **Active** (and optionally **Start Up**) next to FeatureWorks.

## Install

1. **Tools > Macro > New...**, save it as `FeatureWorksStepImporter.swp`
   anywhere convenient.
2. SOLIDWORKS opens the VBA editor. In the Project Explorer, delete the
   empty `Module1` it created (right-click > Remove Module1), then
   right-click the project > **Import File...** and select
   [FeatureWorksStepImporter.bas](FeatureWorksStepImporter.bas).
3. Close the VBA editor and save the macro.

## Add the toolbar button

1. **Tools > Customize**.
2. On the **Commands** tab (or **Macros** tab in older versions), pick the
   **Macro** category, drag the blank button onto any toolbar.
3. When prompted, browse to `FeatureWorksStepImporter.swp` and select the
   `RunFeatureWorksOnStepFiles` macro entry point.
4. Pick an icon/tooltip text if asked, then **OK**.

## Use it

Click the new button. It opens a file picker filtered to `*.step`/`*.stp`;
pick a file, it imports and recognizes it, then asks if you want to do
another. When you're done, click **No** and it shows a summary of how many
files succeeded/failed.

## Notes / assumptions

- A STEP file that imports as an assembly (multiple separate bodies/parts
  rather than one part doc) is skipped — FeatureWorks recognizes features
  within a single part document. Adjust `ImportAndRecognizeStep` if you need
  per-component handling for multi-part STEP files.
- The bitmask constants mirror SOLIDWORKS's `fwAutomaticRecognitionOptions_e`
  and `fwFeatureCreationOptions_e` enums directly (hardcoded, so the macro
  doesn't need a VBA type-library reference to FeatureWorks to run on any
  machine). See the comments at the top of the `.bas` file if you want to
  add sheet-metal types (`fwBaseFlange`, `fwSketchedBend`,
  `fwAutoEdgeFlange`, `fwAutoHemFlange`) back in, or re-include
  `fwVolume`.
- Output overwrites an existing `.SLDPRT` of the same name without
  prompting.
