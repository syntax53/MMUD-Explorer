# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## What this is

MMUD Explorer is a **Visual Basic 6.0** Windows desktop application — a database viewer and
combat/exp calculator for the game MajorMUD. It reads a Microsoft Access `.mdb` game database
via DAO and presents explorers, comparison views, maps, and calculators. Output is a single
compiled `mudexplr.exe`.

## Build, run, and release

There is **no command-line build and no automated test suite**. The project is built in the
**VB6 IDE** (`MMUD_Explorer.vbp`, `Type=Exe`, startup is `Sub Main` in `modMain.bas`).

- Build: open `MMUD_Explorer.vbp` in VB6 → File → Make `mudexplr.exe`.
- The compiled `mudexplr.exe` is committed to the repo and updated as part of a release.
- Dependencies (must be registered on the build machine): DAO 3.6 (`Dao360.dll`), Microsoft
  Scripting Runtime (`scrrun.dll`), `mscomctl.OCX`, `ComDlg32.OCX`, `msstdfmt.dll`, and the
  bundled `exlimiter.ocx`.

### Dev-mode flags — turn OFF before a release build
Four flags are marked `'TURN OFF BEFORE RELEASE - LOC n/4`. All four must be flipped for a
release, and the version-caption string updated:
- `modMain.bas:2` — `#Const DEVELOPMENT_MODE = 1` → `0`
- `frmMain.frm` `bDPIAwareMode = False` (LOC 2/4)
- `frmMain.frm` `bDebugExecTime = False` (LOC 3/4)
- `frmMain.frm` version caption `" vYYMMDDx"` (LOC 4/4) — bump per build
In dev mode, the title shows "(DEV MODE)" and a `_DebugLog.txt` is written.

## File-format rules (important)

- **Line endings must be CRLF** for `.vbp`, `.frm`, `.cls`, `.bas`, `.ini` — enforced by
  `.gitattributes`. Do not let edits convert them to LF.
- **Encoding: `.frm`/`.bas`/`.cls` are Windows-1252 (ANSI), not UTF-8.** Their comments contain
  high-byte characters (em-dashes `—`, curly quotes `’ “ ”`). The built-in **Edit/Write tools
  read and rewrite the whole file as UTF-8, which silently strips every high byte across the
  entire file** — corrupting comment punctuation far from the intended change. No BOM is added, so
  the result stays valid ASCII and still compiles, making the damage easy to miss in a quick diff.
  Do **not** use Edit/Write on these files. Apply changes with a byte-exact **Latin-1 (codepage
  28591) round-trip** in PowerShell instead:
  `$enc=[Text.Encoding]::GetEncoding(28591); $t=[IO.File]::ReadAllText($p,$enc); $t=$t.Replace($old,$new); [IO.File]::WriteAllText($p,$t,$enc)`
  (normalize the search/replace text to CRLF first, guard with an occurrence==1 check). Afterward,
  verify the high-byte count is unchanged vs HEAD (`tr -cd '\200-\377' < file | wc -c`), there are
  zero lone LFs, and a high-byte-stripped diff vs HEAD shows only the intended edits. To recover a
  file already corrupted this way (when its ASCII still matches HEAD), `git checkout HEAD -- <file>`
  then re-apply via the Latin-1 method. (This README/CLAUDE.md and other `.md`/`.txt` files are
  UTF-8/ASCII and edit normally.)
- `.frm` files start with a designer-generated block of control definitions, followed by the
  code. Hand-editing the control/layout block is risky; prefer editing the procedure code at
  the bottom. `frmMain.frm` is ~1.5 MB — read targeted ranges, not the whole file.
- `.frx` files are **binary** form resources paired with each `.frm`. Never edit them by hand.
- `.OBJ` files are stale compiled artifacts; ignore them — the real sources are `.bas`/`.frm`/`.cls`.
- `settings.ini` (user config) and `*.mmec` (saved character files) are gitignored; do not commit them.

## Architecture

**Global state.** `modMain.bas` holds the entry point (`Sub Main`) and a large block of global
variables. The currently loaded character is represented by many `nGlobalChar*` globals (weapon
stats, accuracy, magery, bless spells, etc.) rather than an object — most calculators read these
directly.

**Database layer.** `modMMudDatabase.bas` owns the DAO `Database` (`DB`) and a set of global
`Recordset` objects opened once at load: `tabItems`, `tabMonsters`, `tabSpells`, `tabRooms`,
`tabShops`, `tabRaces`, `tabClasses`, `tabLairs`, etc. Combat/damage result arrays
(`nCharDamageVsMonster()`, `nMonsterDamageVsChar()`, …) also live here.

**Game-engine variants.** MME supports multiple MajorMUD server engines: **stock**,
**GreaterMUD**, and **Paramud**. The data version drives behavior via globals `nGlobalDatVer`,
`nNMRVer`, and the boolean **`bGreaterMUD`** (in `modMMudFunc.bas`). Formula branches throughout
the code switch on these — when changing any game formula, check whether it needs a stock vs.
GreaterMUD/Paramud branch and which data-version threshold gates it. Two recovery facts that affect any regen/rest formula: `nCharHPRegen` is handled in *resting-rate* form (already x3, so passive HP = HPRegen/3 per 30s, rest = HPRegen per 20s), and a character can rest (HP) **or** meditate (MP) but never both and never in combat (passive always ticks).

**Key modules**
- `modMMudFunc.bas`, `modSyntaxsFunc.bas` — core MajorMUD game formulas/helpers.
- `modExpPerHour.bas` — exp/hr prediction models: four selectable models (A/B/C/D); D is a round-by-round sim and the recommended one. See `docs/exp-per-hour-models.md` for game mechanics, the calibration harness, and how to add a model.
- `clsMonsterAttackSim.cls` — monster attack/combat simulation (exposes `bGreaterMUD` property).
- `modItemParse.bas` — item text/detail parsing.
- `modListViewExt.bas` — ListView sorting/grouping used across result tabs.
- `modForms.bas`, `modMenuSubClass.bas`, `modMonitors.bas`, `General.bas` — Win32 subclassing,
  multi-monitor/DPI, and window helpers (lots of `Declare`d API calls).
- `modSettings.bas` — `settings.ini` read/write and working-directory globals.

**Forms** (`frm*.frm`) are the UI: `frmMain` (main multi-tab window), `frmMap`/`frmMapLegend`
(graphical room explorer), `frmResults`, calculators (`frmSwingCalc`, `frmHitCalc`, `frmBSCalc`,
`frmExpCalc`, `frmCoinConvert`, `frmMonsterAttackSim`), `frmPasteChar`/`frmLoadChar` (import
characters), `frmSpellBook`, `frmMegaMUDPathing`, and option/dialog forms.

CLI args: `mudexplr.exe` accepts a `.mdb` (database) and/or `.mmec` (character) file path.

**Dark mode** (`modTheme.bas` owns all of it; **full handoff reference: `docs/darkmode.md`**).
`bDarkMode` is read from `[Settings] DarkMode`
in `Sub Main` (settings path is resolved early there via `bGlobalNewINICreated`); the toggle is
`chkDarkMode` in frmSettings and **requires an app restart** — the design-time colors ARE the
light theme, and dark is applied over them at load. Every form calls `ApplyDarkTheme Me` at the
top of `Form_Load` (a no-op in light mode); **new forms must add this call**. Rules that MUST be
followed when touching UI code:

- Any runtime color assignment goes through `TColor(...)` (general; maps system colors to the
  dark palette, lightness-inverts explicit RGB) or `TBtnColor(...)` (button faces only).
  Inversion is an involution — **never pass a value through `TColor` twice** (note
  `ColorListviewRow` already applies it internally; its callers pass raw colors). Code that
  *compares* a themed control's color must compare against the same `TColor(...)`-wrapped value
  (see `modListViewExt`).
- In dark mode a frame's real caption is not drawn; an overlay label (`lblDkFrameCap*`) stands
  in. Change frame captions via `SetFrameCaption fra, "..."` and caption colors via
  `SetFrameForeColor fra, TColor(...)` — raw `.Caption`/`.ForeColor` assignments silently won't
  display. Also note VB6 fires `chk_Click` only on value *change*, so design-time-default states
  need explicit color init in `Form_Load` (see the `chkGlobalFilter` block).
- Opt-outs: put `notheme` in a control's `Tag`, and any control with a custom (non-system)
  opaque `BackColor` is skipped automatically — that's what protects the map room cells
  (`lblRoomCell`), the black char-stat panel, the map-legend swatches, and frmMap's options
  panel. Colors on those are data/intentional styling, not chrome.
- All CommandButtons are `Style=1 'Graphical` so `BackColor` works (renders identically in
  light mode since the app has no comctl6 manifest); **new buttons should be Style=1 too**.
  VB6 button captions are always black — hence the mid-gray `DK_BTN_FACE`, not a true dark face.
- ComboBox and Frame borders are overdrawn from comctl `SetWindowSubclass` post-paint procs
  (gated by `gbAllowSubclassing`; frame border geometry is packed into `dwRefData`). frmMain's
  menu bar is owner-drawn dark via the `WM_UAH*` messages in `modMenuSubClass`. VB6 gotchas that
  bit here: `AddressOf Foo` cannot be used inside `Foo` itself, and `Not (x And &H80000000)` is
  bitwise (truthy either way) — compare with `= 0` instead.
- Intentionally still system-light: scrollbars, ListView column headers, popup/context menus,
  MsgBoxes (ListView gridlines are simply turned off in dark). **Do not** retry the uxtheme
  dark-mode ordinals (`SetPreferredAppMode`/`FlushMenuThemes`/`AllowDarkModeForWindow`/
  `SetWindowTheme "DarkMode_Explorer"`) — they act process-wide, did nothing useful here, and
  destabilized the VB6 IDE (run-time error 97 at IDE close).
- Button glyphs: pictures are embedded in binary `.frx` (re-assign via the IDE only); the loose
  source images live in the repo root. VB6 transparency is 1-bit — hard-edged glyphs, or
  antialiasing pre-blended toward `#C4C4C4` (midpoint of the light `#F0F0F0` and dark `#989898`
  button faces); 24-bit BMP + `UseMaskColor`/magenta beats GIF for color depth.

## Conventions

- `Option Explicit` is used throughout; keep it. Hungarian-style prefixes are pervasive
  (`b`=Boolean, `n`=numeric, `s`=String, `tab`=Recordset, `frm`=Form, `mod`=Module, `cls`=Class).
- The changelog lives in `docs/changelog.txt` (current) and `docs/changelog.archive.txt`;
  user-facing change notes are also mirrored at the top of `README.md`. Update these when
  shipping user-visible changes.

## Git

`master` is the main branch; work happens on dated `dev-*` branches. Source-of-truth changes are
the `.bas`/`.frm`/`.cls` files; a release commit also includes the rebuilt `mudexplr.exe`.
