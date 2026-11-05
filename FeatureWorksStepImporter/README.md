# FeatureWorks STEP Importer (VBA macro)

Two SOLIDWORKS VBA macros for automating STEP → native SOLIDWORKS conversion:

- **`FeatureWorksStepImporter.bas`** — batch-converts multiple independent
  single-part STEP files, one at a time.
- **`FeatureWorksStepAssemblyImporter.bas`** — converts one STEP file (part
  *or* assembly) into a native SOLIDWORKS assembly with fully-recognized
  native parts, with an output folder you choose.

Both macros run FeatureWorks so the saved result shows real recognized
features (`Extrude1`, `Hole1`, `Fillet1`, ...) instead of a single dumb
`Imported1` body.

## How recognition is actually driven (read this first)

The "official" documented way to automate this —
`IFeatureWorksApp::RecognizeFeatureAutomatic` + `CreateFeatures` — is a
known dead end: it returns 0 recognized features for a lot of people even
run exactly per SOLIDWORKS's own sample code. Multiple independent forum
reports hit the exact same symptom this project hit during testing:
([one](https://3dswym.3dexperience.3ds.com/question/solidworks-user-forum/hello-i-want-to-make-a-macro-for-automatic-feature-recognition-of-an-imported-model-and-create-the-features-under-the-tree-i-have-tried-the-below-codes-it-s-running-but-didn-t-get-any-feature_Yr2EVlLET-WtVu6-4K6IDA),
[two](https://3dswym.3dexperience.3ds.com/question/solidworks-user-forum/need-c-example-of-how-use-recognizefeatureautomatic-with-api_KkmjBBw-SdKn5tLjFVli5A)).

Both macros instead drive FeatureWorks through
`ISldWorks::CallBack("fworks@FWPlaybackManager", 0, "<command>")` — an
**undocumented** internal interface, but the exact sequence SOLIDWORKS's
own macro recorder captures when you manually run **Insert > FeatureWorks >
Recognize Features** with Automatic selected in the PropertyManager. This
project's sequence was captured and verified working that way (selecting
the imported body feature first is required, regardless of recognition
mode — none of the official samples do this, which is almost certainly why
they don't work for people):

```vb
feat.Select2 False, -1   ' the imported body feature - required
swApp.CallBack "fworks@FWPlaybackManager", 0, "SetAutomaticFeatureRecognition 0"
swApp.CallBack "fworks@FWPlaybackManager", 0, "SetInteractiveFeatureRecognition 1"
swApp.CallBack "fworks@FWPlaybackManager", 0, "RecognizeFeatures"
swApp.CallBack "fworks@FWPlaybackManager", 0, "SetAutoDimOptions 0 0 0 0"
swApp.CallBack "fworks@FWPlaybackManager", 0, "SetAutoRelOptions 0"
```

Yes, the flag names look backwards (`SetAutomaticFeatureRecognition 0` /
`SetInteractiveFeatureRecognition 1` for *Automatic* mode) — that's exactly
what the recorder captured with Automatic selected in the UI, so it's used
verbatim rather than "corrected" based on what the names seem to imply.

### One-time manual setup — WHICH feature types get recognized

Neither the documented API nor the `FWPlaybackManager` commands above
include anything for choosing *which* feature types to recognize (Standard
features vs. sheet metal, Extrude/Revolve/Holes/Fillets/Ribs individually,
Volume on/off). That's because the FeatureWorks PropertyManager just
**remembers its own settings** between runs, the same way most SOLIDWORKS
PropertyManagers do — so there's nothing for a macro to set.

**Before running either macro for real work**, do this once by hand:

1. Open any STEP file.
2. **Insert > FeatureWorks > Recognize Features**.
3. In the PropertyManager, select **Automatic**, **Standard features**
   (not sheet metal), and check/uncheck exactly the feature types you want
   (uncheck **Volume** if you don't want volume-based recognition; Draft
   isn't offered in Automatic mode at all, so there's nothing to uncheck
   there).
4. Click the green checkmark to finish.

SOLIDWORKS keeps that configuration as the default for every later
`RecognizeFeatures` call — including the ones these macros trigger — until
you change it again by hand the same way.

## Prerequisites

- FeatureWorks must be licensed and enabled: **Tools > Add-ins**, tick
  **Active** (and optionally **Start Up**) next to FeatureWorks.
- Do the one-time manual setup above at least once.

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

A file "succeeds" only if no `Imported*` feature remains in its tree
afterward — if FeatureWorks only partially recognized the body, that file
is reported as failed (with the save path) so you can open it and check
what's still dumb geometry, rather than silently shipping a half-recognized
part.

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

A summary MsgBox reports how many components were processed, how many of
those came out "clean" (no `Imported*` feature left — same success bar as
the single-part macro), how many were skipped, and the mate-repair result.

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

- Both macros assume the imported body's feature is literally the *first*
  feature in a freshly-opened/freshly-externalized part (true for a dumb
  solid nothing else has touched yet) and select it with
  `FirstFeature.Select2` rather than matching on the name `"Imported1"` —
  more robust across locales/renames. Multi-body imports (several
  `ImportedN` features) only get the first one selected; extend
  `RunFeatureWorksRecognition` if you need all bodies selected at once.
- "Recognized cleanly" is judged purely by whether any feature named
  `Imported*` remains — it doesn't verify which specific feature *types*
  came out, since that's governed by the one-time manual setup above, not
  by anything the macro controls or can inspect reliably.
- Output overwrites an existing file of the same name without prompting.
