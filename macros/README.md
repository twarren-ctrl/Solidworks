# Tank + Dual Screw Conveyor Macro

This folder contains a SolidWorks VBA macro that builds a tank in **inches** from the provided profile and adds two screw-conveyor placeholders and motor placeholders.

## File
- `tank_dual_screw_macro.vba`

## What it creates
1. Side-profile tank body extruded to **26 in width**.
2. Shell feature with **0.25 in thickness**.
3. Two straight screw-conveyor placeholders aligned to transfer/discharge paths.
4. Two cylindrical motor placeholders at conveyor ends.

## Setup
1. Open the VBA macro in SolidWorks.
2. Update `TEMPLATE_PATH` to your local `Part.prtdot` template path (optional but recommended).
3. Run `main`.

## Notes
- If `TEMPLATE_PATH` is invalid, the macro falls back to `swApp.NewPart` (SolidWorks default part behavior), then makes a final default-template preference lookup attempt.
- The conveyor solids are simplified placeholders (outer screw OD + shaft). If you want true helical flights, replace `BuildStraightScrew` with a helix+sweep feature workflow.
- All dimensions are at the top of the macro for quick edits.

- Shell creation is attempted with cross-version API fallback (`InsertFeatureShell` then `InsertShell`).
