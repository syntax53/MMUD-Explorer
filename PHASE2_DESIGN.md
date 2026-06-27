# Phase 2 — Dynamic quest UI, custom quests, stable keys, NMR, dark mode

Phase 1 made the *stat math* data-driven while keeping the fixed 12-checkbox /
4-combo panel. Phase 2 makes the *panel itself* dynamic, adds player/realm custom
quests, moves persistence onto stable keys, lets NMR write a per-realm file, and
skins the whole app in dark mode.

This document is the plan. The foundation pieces (data model, theme engine) are
delivered and verified now; the form surgery + new editor form + NMR writer are
the next implementation step (they touch the 42k-line `frmMain.frm` and a new
form, so they ship as reviewable patches, since none of it can be compiled in
this environment — VB6 needs the Windows IDE).

## Status in this drop

| Piece | Status |
|---|---|
| `modQuestConfig.bas` v2 — dynamic list, stable keys, labels, import/export rules, custom load/override/append, custom save, summary helper | **Done, verified** |
| `MME-Quests.txt` v2 — keyed, labeled, with import/export + custom template | **Done, verified** |
| `modTheme.bas` — dark-mode engine (palette, recursive control theming, DWM dark title bar, dark combo/list theming, persisted toggle) | **Done** |
| Dark-mode skin design (mockup) | **Done** |
| `frmMain.frm` surgery — scrollable list container, runtime row generator, applier rewrite, INI/paste on keys, theme call | **Next** |
| `frmQuestEdit.frm` — add/edit a custom quest | **Next** |
| NMR `frmMME_Export.frm` — write per-realm `MME-Quests.txt` | **Next** |

## 1. Data model (delivered)

A quest is no longer a 0–11 slot; it's an entry in an ordered list with a stable
`Key`. `modQuestConfig` v2 exposes `g_Quests()` / `g_QuestCount`, loads
`MME-Quests.txt` (built-ins) then `MME-QuestsCustom.txt` (override-by-key or
append-new), and can save edited/custom quests back out. Verified: the 12
built-ins parse in order with correct keys, the 4 choice quests keep their option
counts (6/7/4/3), import/export rules are captured, and a custom file correctly
overrides a built-in *and* appends a new quest. The Phase-1 numeric-delta
equivalence still holds for all 34 quest/engine/option combinations.

File format v2 adds, per quest: `Key`, optional `Import = abil:min`, optional
`Export = abil|val`, and per-option `Option<k> = <label> | <terms>`. Custom
quests use a slug section id, e.g. `[Quest soulforge_trial]`.

## 2. Dynamic list UI (next)

The "Completed Quests" frame `fraChar(4)` becomes a scrollable list:

* A clipping `picQuestView` (PictureBox) fills the frame; an inner `picQuestRows`
  holds the rows and is scrolled by offsetting its `.Top`; a `vsQuests`
  vertical scrollbar drives it. (The old fixed checkboxes/combos/labels are
  removed from the frame; one hidden template member of `chkCharQuests` and
  `cmbCharQuestOpts` stays inside `picQuestRows` so the control arrays can be
  `Load()`-cloned at runtime into that container.)
* `modQuestUI.BuildQuestList` iterates `g_Quests()` and, per quest, `Load`s a
  checkbox (caption = `Name` + summary from `SummaryForTerms`); choice quests
  also get a `Load`ed combo populated from `OptionLabel`. It records mapping
  arrays `mChkIdx(i)` / `mCmbIdx(i)` / `mKey(i)` so the rest of the code finds a
  quest's controls by list position or key.
* A small toolbar (Add / Edit / Remove / Realm file…) drives custom-quest CRUD.

This keeps every quest reference funneled through the mapping + keys instead of
hard-coded 0–11, which is what makes custom quests and reordering safe.

## 3. Stat applier (next — small change)

The Phase-1 `ApplyQuestRewards` loop changes from `For qi = 0 To 11` /
`chkCharQuests(qi)` to iterating `g_Quests()` and reading the mapped control
(`chkCharQuests(mChkIdx(i))`, `cmbCharQuestOpts(mCmbIdx(i))`). The per-term
application (`ApplyOneQuestTerm`) and the two-pass encum/str ordering are
unchanged, so the verified numeric behavior carries over.

## 4. Persistence on stable keys (next)

* **INI save/load** (`Quest0..11`, `Quest_2nd`, …) → key-based
  (`Quest_<key>=1`, `QuestOpt_<key>=<idx>`). A legacy reader still understands the
  old numeric keys via `QuestIndexBySlot`, so existing saved characters migrate.
* **Paste import / export** become data-driven via the `Import` / `Export` rules
  in the file. The form's hard-coded ability `Select Case` is replaced by a loop
  over `g_Quests()` that applies each quest's rule, gated by the quest's engine
  (reproducing the old `If bGreaterMUD` wrapper). Custom quests can opt in by
  declaring their own `Import`/`Export`.

  **Flag for review:** the current import code maps ability **203 → the
  Cartographer checkbox** and **202 → the Loremaster checkbox**, which is the
  opposite of its own inline comments (which say 202 = Cartographer, 203 =
  Loremaster). The v2 default file preserves the *current code behavior*
  (`cartographer Import = 203`, `loremaster Import = 202`). If that was a latent
  bug, swap the two `Import` lines in `MME-Quests.txt` — no recompile needed.

## 5. NMR per-realm file (next)

`Nightmare-Redux` already has `frmMME_Export.frm` that bridges realm data to MME.
Phase 2 adds an option there to emit a realm-specific `MME-Quests.txt` next to the
exported data, so a custom realm ships its quest rewards alongside its MDB. NMR
writes the same v2 format MME reads; no shared code, just an agreed file.

## 6. Dark mode (delivered engine + design)

`modTheme.bas` centralizes the palette and applies it per control type. Because
VB6 has no native dark mode it: sets BackColor/ForeColor where the control
supports it; uses DWM (`DWMWA_USE_IMMERSIVE_DARK_MODE`) to darken the window
title bar on Win10 1809+; and uses `SetWindowTheme("DarkMode_Explorer")` to
darken combo dropdowns, list boxes, list/tree views and scroll bars where the OS
allows. Call `ApplyTheme(Me)` at the end of each form's `Form_Load` (and after
`BuildQuestList` so generated rows get themed); `ToggleDarkMode(Me)` flips and
persists it. The palette harmonizes with the already-black `frmLoad` launcher
splash.

Two honest VB6 limitations: standard `CommandButton`s only accept BackColor when
`Style = 1 (Graphical)` — buttons left default stay system-gray; and the classic
scrollbar can't be fully recolored, so the list uses a thin custom-drawn scroll
track (as in the mockup) rather than the OS scrollbar.

## Integration order / caution

`modQuestConfig.bas` v2 **supersedes** the Phase-1 module (it renames
`g_QuestDefs` → `g_Quests` and generalizes the type). Do **not** drop the v2
module onto a Phase-1-patched `frmMain.frm` on its own — the Phase-1 applier
references the old names. The Phase 2 form surgery (next) updates the applier to
the v2 contract; ship them together. `modTheme.bas` is independent and can be
added any time.
