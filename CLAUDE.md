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
GreaterMUD/Paramud branch and which data-version threshold gates it.

**Key modules**
- `modMMudFunc.bas`, `modSyntaxsFunc.bas` — core MajorMUD game formulas/helpers.
- `modExpPerHour.bas` — exp/hr prediction models (modelA/B/C selectable in settings).
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

## Conventions

- `Option Explicit` is used throughout; keep it. Hungarian-style prefixes are pervasive
  (`b`=Boolean, `n`=numeric, `s`=String, `tab`=Recordset, `frm`=Form, `mod`=Module, `cls`=Class).
- The changelog lives in `docs/changelog.txt` (current) and `docs/changelog.archive.txt`;
  user-facing change notes are also mirrored at the top of `README.md`. Update these when
  shipping user-visible changes.

## Git

`master` is the main branch; work happens on dated `dev-*` branches. Source-of-truth changes are
the `.bas`/`.frm`/`.cls` files; a release commit also includes the rebuilt `mudexplr.exe`.
