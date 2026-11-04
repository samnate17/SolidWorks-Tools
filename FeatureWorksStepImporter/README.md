# FeatureWorks STEP Importer (VBA macro)

Two SOLIDWORKS VBA macros for automating STEP → native SOLIDWORKS conversion:

- **`FeatureWorksStepImporter.bas`** — batch-converts multiple independent
  single-part STEP files, one at a time.
- **`FeatureWorksStepAssemblyImporter.bas`** — converts one STEP file (part
  *or* assembly) into a native SOLIDWORKS assembly with fully-recognized
  native parts, with an output folder you choose.

Both run FeatureWorks feature recognition with the same settings:

- **Recognition mode:** Automatic (`RecognizeFeatureAutomatic`)
- **Feature type:** Standard features only (not sheet metal)
- **Auto-recognized types:** Extrude/boss/cut, Revolve, Holes, Fillets &
  Chamfers, Ribs — i.e. everything available in automatic mode **except
  Volume**. Draft isn't included either, but that's automatic for free:
  FeatureWorks has no automatic-mode option for Draft at all — it's only
  recognized interactively.

- **Recognition mode:** Automatic (`RecognizeFeatureAutomatic`)
- **Feature type:** Standard features only (not sheet metal)
- **Auto-recognized types:** Extrude/boss/cut, Revolve, Holes, Fillets &
  Chamfers, Ribs — i.e. everything available in automatic mode **except
  Volume**. Draft isn't included either, but that's automatic for free:
  FeatureWorks has no automatic-mode option for Draft at all — it's only
  recognized interactively.

## Prerequisites

- FeatureWorks must be licensed and enabled: **Tools > Add-ins**, tick
  **Active** (and optionally **Start Up**) next to FeatureWorks.

## Install (both macros live in one macro project)

1. **Tools > Macro > New...**, save it as `FeatureWorksStepImporter.swp`
   anywhere convenient.
2. SOLIDWORKS opens the VBA editor. In the Project Explorer, delete the
   empty `Module1` it created (right-click > Remove Module1), then
   right-click the project > **Import File...** twice, once for each
   [FeatureWorksStepImporter.bas](FeatureWorksStepImporter.bas) and
   [FeatureWorksStepAssemblyImporter.bas](FeatureWorksStepAssemblyImporter.bas).
3. Close the VBA editor and save the macro.

## Add toolbar button(s)

1. **Tools > Customize**.
2. On the **Commands** tab (or **Macros** tab in older versions), pick the
   **Macro** category, drag a blank button onto any toolbar.
3. When prompted, browse to `FeatureWorksStepImporter.swp` and pick either
   `RunFeatureWorksOnStepFiles` (batch single parts) or
   `RunFeatureWorksOnStepAssembly` (assembly conversion) as the entry point.
   Repeat to add both as separate buttons.
4. Pick an icon/tooltip text if asked, then **OK**.

---

## `RunFeatureWorksOnStepFiles` — batch single parts

Click the button. It opens a file picker filtered to `*.step`/`*.stp`; pick
a file, it imports and recognizes it, saves it as a native `.SLDPRT` **next
to the source STEP file** (same name, `.SLDPRT` extension), closes it, then
asks if you want to do another. When you're done, click **No** and it shows
a summary of how many files succeeded/failed.

A STEP file that imports as an assembly is skipped by this macro — use
`RunFeatureWorksOnStepAssembly` for those.

---

## `RunFeatureWorksOnStepAssembly` — STEP → native assembly + parts

Click the button, pick **one** STEP file (part or assembly), then pick an
**output folder**. What happens next depends on what the STEP file is:

- **Single part:** recognized directly, saved into the output folder.
- **Assembly:** for every top-level component, the macro:
  1. Saves that component's own document out to its own `.SLDPRT` in the
     output folder (if the component came in "virtual"/embedded — which is
     typical for a STEP-assembly import — this `SaveAs` on its own document
     is what gives it a real external file; SOLIDWORKS does this under the
     hood whenever you "Make Independent" a virtual component by hand).
  2. Runs FeatureWorks on that component **in place** — since the component
     still references that exact file (we never rename/relink it to a
     different file), there's no separate "replace component" step. This
     also gives mates the best possible chance of surviving, since nothing
     about the component's identity changed, only its geometry's feature
     history.
  3. Rebuilds the assembly.
  4. On **SOLIDWORKS 2025 SP01 or newer only**, runs
     `AssemblyDoc::AutoMateRepair` to auto-heal any mates broken by the
     dumb-solid → recognized-features swap. On older versions this step is
     skipped (reported in the summary) — check the Mates folder in the
     FeatureManager tree by hand for anything greyed-out or red.
  5. Saves the assembly into the output folder as `<name>.SLDASM`.

A summary MsgBox reports how many components were recognized, how many were
skipped, and the mate-repair result.

### Known limitations / things to verify on your end

I don't have a SOLIDWORKS install to test this against, so a few points are
my best-grounded design decisions rather than confirmed behavior — worth
checking the first time you run it, ideally on a throwaway copy:

- **Nested sub-assemblies are not recursed into.** Only top-level components
  are processed; a component that is itself a sub-assembly is skipped and
  counted in "skipped." Extend `ProcessOneComponent` in the `.bas` file if
  you need that.
- **Virtual-component externalization.** The step that calls `SaveAs` on a
  virtual component's document to give it its own file is the part most
  worth double-checking — if a component doesn't externalize cleanly, the
  known manual fallback is to select all components in the assembly,
  right-click, and **Make Independent** before running the macro.
- **Mate healing is best-effort, not guaranteed**, and only runs at all on
  2025 SP01+. No CAD system can promise mates survive an arbitrary geometry
  swap — faces a mate references by persistent ID can legitimately not
  exist anymore once FeatureWorks rebuilds the body via a different feature
  history. Always check the Mates folder after running this.
- The macro leaves the assembly and its component documents open afterward
  (rather than closing them) so you can inspect the result immediately.
- Output filenames strip SOLIDWORKS's `-N` instance suffix and `^AssemblyName`
  virtual-component suffix from each component's display name; name
  collisions get a `_2`, `_3`, ... suffix.

## Notes / assumptions (both macros)

- The bitmask constants mirror SOLIDWORKS's `fwAutomaticRecognitionOptions_e`
  and `fwFeatureCreationOptions_e` enums directly (hardcoded in both `.bas`
  files, so neither macro needs a VBA type-library reference to FeatureWorks
  to run on any machine). See the comments at the top of either file if you
  want to add sheet-metal types (`fwBaseFlange`, `fwSketchedBend`,
  `fwAutoEdgeFlange`, `fwAutoHemFlange`) back in, or re-include `fwVolume`.
- Output overwrites an existing file of the same name without prompting.
