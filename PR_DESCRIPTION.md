# PR: Data-driven, custom-realm-extensible "Completed Quests" system

## Summary
Moves the hardcoded **Completed Quests** stat bonuses and the NMR character
ability-code export/import out of `frmMain` and into an external, per-realm data
file. Custom realms can add their own quests — and the new ability codes they
mint to track them — without recompiling MME. The fixed 12-checkbox / 4-combo
panel becomes a dynamic, scrollable list that grows with the data file.

## What changed
1. **Data-driven stat math.** The hardcoded reward blocks are replaced by a
   reward-term engine that reads `MME-Quests.txt` (with per-realm overrides in
   `MME-QuestsCustom.txt`). Behavior is preserved exactly for the stock realm.
2. **Dynamic scrollable quest UI.** Built-in quests keep their original control
   indices (so all legacy INI/paste/export references keep working); custom
   quests load into new rows. Dark-mode aware if the dark-mode PR is also merged.
3. **Key-based character persistence.** Selections save under a stable
   `CompletedQuests` key (quest-key, not slot), with verbatim legacy-slot
   fallback so existing character files still load.
4. **Data-driven NMR export/import.** Per-quest `Export`/`Import` rules and
   per-option `Option<k>_Export` codes drive the clipboard ability-code round-trip,
   so custom realms export/import their own tracking abilities. The genuinely
   stock-specific glue is preserved verbatim and *not* re-encoded into the format:
   alignment-tier codes (126/127/128, keyed to global alignment), the
   class-derived Dread import (ability 221), and the cumulative Renfry conquest
   import (208/209).

## Files
- `frmMain.frm` — quest panel rebuilt; reward/export/import made data-driven.
- `modQuestConfig.bas` — **new**; quest model, file parser, export/import helpers,
  embedded default (kept in sync with `MME-Quests.txt`).
- `MME-Quests.txt` — **new**; shipped default quest definitions. Ships next to the
  exe and is read at runtime; `MME-QuestsCustom.txt` (same format) is optional.
- `MMUD_Explorer.vbp` — adds the `modQuestConfig` module.
- `README_configurable_quests.md`, `PHASE2_DESIGN.md` — format + design notes.

## Apply / build
- `git apply custom_quests.patch` on current `main`, **or** drop in the full files.
- Add `modQuestConfig` to the project (already in the included `.vbp`).
- Place `MME-Quests.txt` beside `mudexplr.exe`.
- Build in the VB6 IDE → Make `mudexplr.exe`.

## Verification
- Stat math reproduces the stock output for every quest/option/engine combination.
- Export reproduces the stock ability-code set for every quest/option/engine;
  import reproduces stock thresholds and engine scope. (Ability order within the
  `ABILS:` line is a set and may differ; it is functionally identical.)
- One intentional, strictly-more-correct change: an empty `ABILS:` line that stock
  could emit in an exotic edge case is no longer written.

## Notes / follow-ups
- The Add/Edit/Remove/Realm toolbar buttons are present but stubbed pending a
  small `frmQuestEdit` dialog (separate change).
- A pre-existing mislabel in stock comments (Cartographer vs Loremaster ability
  labels) is preserved; the actual codes are faithful to current behavior.
