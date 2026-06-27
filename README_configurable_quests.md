# Data-driven "Completed Quests" stat bonuses (Phase 1)

Makes the hard-coded Completed-Quest stat bonuses in the Char tab read from an
external, editable text file instead of a baked-in `Select Case`. Custom realms
can now adjust quest rewards by editing a text file — no recompile required.

This is **Phase 1**: the panel layout, checkboxes, and option dropdowns are
unchanged. Only the *stat application* becomes data-driven. (UI revamp,
per-realm file writing from the editor, and stable-key import/paste are deferred
to Phase 2.)

## Files in this contribution

| File | What it is |
|---|---|
| `modQuestConfig.bas` | **New module.** Types, file loader/parser, term parser, and an embedded default used to recreate the file if it's missing. |
| `MME-Quests.txt` | **New data file.** The shipped default reward set (also auto-recreated at runtime if absent). Drop next to the EXE. |
| `frmMain.frm` | Two hard-coded quest blocks replaced by `ApplyQuestRewards(...)` calls; two new private subs (`ApplyQuestRewards`, `ApplyOneQuestTerm`) added; one `Call LoadQuestDefs` added in `Form_Load`. |
| `MMUD_Explorer.vbp` | One line added to register the new module. |

You can apply this two ways:
* **Patch:** `git apply CONTRIB.patch` from the repo root (covers all four files).
* **Drop-in:** copy `modQuestConfig.bas`, `MME-Quests.txt`, the modified
  `frmMain.frm`, and `MMUD_Explorer.vbp` over the originals.

## How it works

* On `Form_Load`, `LoadQuestDefs` reads `App.Path\MME-Quests.txt`. If it's
  missing, it's recreated from the built-in default, so behaviour is identical
  out of the box.
* An optional `MME-QuestsCustom.txt` next to the EXE overrides any `[Quest N]`
  blocks present in it (per-realm tweaks without touching the shipped file).
* `CalcCharacterStats` calls `ApplyQuestRewards` in two passes:
  * `QPASS_PREENCUM` (inside the existing GreaterMUD gate) applies only the
    `encum`/`str` terms, before encumbrance is computed — same ordering as today.
  * `QPASS_MAIN` applies everything else.

### File format (per quest)

```
[Quest N]                 N = panel slot index 0..11
Name   = <text>           shown in stat tooltips: "Quest: <Name> (value)"
Engine = both|stock|gmud  which engine(s) the quest applies under
Choice = <comboIndex>     only for multi-option quests (cmbCharQuestOpts index)
Reward = <terms>          simple quest
Option<k> = <terms>       multi-option quest, k = combo ListIndex
```

Terms are comma-separated `target:value`:

| target | effect |
|---|---|
| `sN` | `lblInvenCharStat(N)` += value (additive stat slot) |
| `accy` | accuracy ability bonus, cumulative |
| `accymax` | accuracy ability bonus, highest-wins (stock Bishop rule) |
| `dodge` | `nGlobalCharPlusDodge` += value |
| `str` | main-stat Strength bonus (`AdjMainStatBonus`) |

A term can carry an engine qualifier: `[stock]sN:v` or `[gmud]sN:v` (no prefix =
both). Slot reference: `2`=AC `4`=Encum `5`=MaxHP `6`=Mana `7`=Crits `9`=SC
`11`=MaxDmg `14`=BS-Min `15`=BS-Max `17`=ManaRegen `19`=Stealth.

## Faithfulness / verification

`verify_quests.py` parses the **original** `Select Case` straight out of
`frmMain.frm`, derives the stat deltas for every quest × engine × option, and
compares them to what `MME-Quests.txt` produces.

```
python verify_quests.py path\to\frmMain.frm path\to\MME-Quests.txt
```

Result: **all 34 quest/engine/option combinations match, 0 differences**
(captured in `verify_output.txt`).

One intentional cosmetic change: in the original, the 6th-Alignment option-1
+50 MaxHP tooltip was copy-pasted as `Quest: Opaline (50)`. Tooltips are now
generated from the quest `Name`, so that line correctly reads
`Quest: 6th Align (50)`. No stat value changes.

## Building (VB6 IDE)

This is a VB6 project; it must be compiled with the VB6 IDE on Windows (there is
no command-line/cross-platform VB6 compiler).

1. Apply the patch or copy the files in.
2. Open `MMUD_Explorer.vbp` in VB6. Confirm `modQuestConfig.bas` appears under
   Modules (the `.vbp` line registers it; if you added files manually, use
   Project → Add Module → Existing).
3. Make sure `MME-Quests.txt` sits next to the built EXE (it auto-creates if
   not, but shipping it is cleaner).
4. File → Make `MMUD Explorer.exe`.

## Submitting to the repo (PR flow)

1. Fork the repo, then `git clone` your fork.
2. `git checkout -b configurable-quest-bonuses`
3. Apply the patch (`git apply CONTRIB.patch`) or copy the files in; `git add -A`.
4. Commit (suggested message):
   ```
   Make Completed-Quest stat bonuses data-driven (MME-Quests.txt)

   Replace the hard-coded quest Select Case in CalcCharacterStats with a
   loader (modQuestConfig) that reads reward definitions from MME-Quests.txt,
   with an optional MME-QuestsCustom.txt override. Panel/UI unchanged.
   Verified to reproduce the original stat deltas for all quest/engine/option
   combinations.
   ```
5. `git push` and open a Pull Request. Paste the verification summary
   (`verify_output.txt`) into the PR description — it shows the change is
   behaviour-preserving.
