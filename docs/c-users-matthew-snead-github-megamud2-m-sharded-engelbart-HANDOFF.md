# Handoff — "Mark MegaMUD Rooms" map feature

**Repo:** C:\Users\matthew_snead\GitHub\MMUD-Explorer (VB6 / MMUD Explorer)
**Branch:** dev-260530
**Plan file:** `c-users-matthew-snead-github-megamud2-m-sharded-engelbart.md` (same folder)
**Status:** Implemented; compiles-by-inspection only. User has visually tested two drawing iterations and approved the final **cell-recolor** approach. **Not yet built in the VB6 IDE or fully runtime-tested.**

## What the feature does

Repurposes the seldom-used `chkMapOptions(12)` checkbox (was "Don't Follow Restricted",
now **"Mark MegaMUD Rooms"**) on both frmMain and frmMap. When enabled it parses MegaMUD
`rooms.md` file(s) into an in-memory dictionary and, while drawing a map, recolors any room
whose computed hash is present:

- **Known + has up/down/both exit** → cell BackColor **bright red** `&HFF&` (overrides the usual up/down colors)
- **Known, no up/down** → cell BackColor **bright white** `&HFFFFFF`
- **Not known** → unchanged (gray `&HC0C0C0` / green up `&HFF00&` / yellow down `&HFFFF&` / cyan both `&HFFFF00`)

Matched rooms also get `MegaMUD: [group] code - name` appended to the cell tooltip.

Data is **in-memory per session** (not persisted, not cached to disk). The checkbox starts
unchecked each session; enabling it triggers a file prompt (or a "refresh?" prompt if data is
already loaded). The old "Don't Follow Restricted" behavior (skip exit types 13–15) was dropped.

## Key technical facts

- **Match key** = `Get_MegaMUD_RoomHash("",Map,Room)` (3 hex) `& Get_MegaMUD_ExitsCode(Map,Room)` (5 hex)
  = the 8-char value in `rooms.md` field 0 (e.g. `98040005`). Same computation `cmdMapMegaRoomFind` uses.
- **rooms.md line** split on `:` → `sArr(0)`=hash, `sArr(5)`=code, `sArr(6)`=group, `sArr(7..)`=name.
- Dict value stored as `"code|group|name"`; tooltip splits on `|`.
- Detection runs **before** the cell-color block in `MapMapExits`; the hash funcs re-seek
  `tabRooms` to the same (Map,Room), so the subsequent `U`/`D` reads stay correct.

## Files changed (all committed-worthy; nothing committed yet)

- **modMMudDatabase.bas** — globals: `dictMegaRooms As Scripting.Dictionary`, `sMegaRoomsRootFile As String`.
- **modMain.bas** — `MegaRooms_EnsurePopulated(bForceRepick)`, `MegaRooms_ScanFolderRecursive`,
  `MegaRooms_ParseFile` (added right after `Get_MegaMUD_RoomHash`). File picker reuses
  `frmMain.oComDag`; recursive scan of all subfolders for `rooms.md`.
- **frmMain.frm** — caption→"Mark MegaMUD Rooms"; removed restricted-skip line; new
  `chkMapOptions_Click` (Index 12: populate-on-demand + redraw, resets checkbox on cancel);
  `MapMapExits` detection+recolor block + tooltip append; `bMegaKnown` added to its Dim line.
- **frmMap.frm** — same as frmMain, plus removed `ExMapNoRestricted` load/save; Index 12 branch
  added to the existing `chkMapOptions_Click`.
- **docs/changelog.txt** + **README.md** — v2.2.1 user-facing note.

Note: an earlier "gap-frame" drawing approach (a `drMegaKnown=14` enum + `Case 14` in
`MapDrawOnRoom`) was tried and **fully removed** — it was invisible on the black background and
conflicted with overlays. Don't re-add it.

## ⚠️ Editing rule (critical)

`.frm`/`.bas` are CP1252. **Do not use Edit/Write** (strips high bytes file-wide). Use the
Latin-1 (codepage 28591) PowerShell round-trip with CRLF normalization and an occurrence==1
guard (see `[[edit-tool-strips-high-bytes]]` memory / CLAUDE.md). After editing, verify high-byte
count unchanged vs HEAD and zero lone LFs:
`for f in frmMain.frm frmMap.frm modMain.bas modMMudDatabase.bas; do tr -cd '\200-\377' < $f | wc -c; done`
Current expected high-byte counts vs HEAD: modMMudDatabase=18, modMain=1, frmMain=6, frmMap=0.
(README.md is LF-only; changelog.txt is CRLF.)

## How to verify (no CLI build exists)

1. Open `MMUD_Explorer.vbp` in the VB6 IDE → File → Make `mudexplr.exe`. Fix any compile errors.
2. Load a known `.mdb`; draw a map for an area present in a `rooms.md`.
3. Check "Mark MegaMUD Rooms" → file prompt → pick the main `Default\rooms.md`. Confirm:
   known rooms turn white (or red if they have up/down), unknown rooms keep normal colors,
   tooltip shows the MegaMUD code/group/name.
4. Toggle off → redraw clears it. Toggle on again → "refresh?" prompt (No keeps data).
5. Repeat on frmMap at a normal and a zoom (`picZoomMap`) size.
6. Confirm restricted exits (class/race/level) are now always followed on both forms.

Sample rooms.md for testing: `C:\Users\matthew_snead\GitHub\MegaMud2\MegaMudSource\Release\Default\ROOMS.md`

## Open / optional items (none blocking)

1. **Release prep**: before a release build, flip the 4 dev-mode flags OFF and bump the version
   caption (see CLAUDE.md "Dev-mode flags"). Set the v2.2.1 date in changelog.txt + README.md.
2. **Safe-refresh** (minor): `MegaRooms_EnsurePopulated` rebuilds the dict before validating row
   count, so refreshing with a bad/empty file discards prior data. Could scan into a temp dict
   and only swap on success.
3. **Performance**: with the option on, each drawn room does 2 extra recordset seeks (plus item
   lookups only for rooms with placed items). Fine for normal maps; could lag on huge zoom maps.
   Optional optimization: hash inline off the already-seeked `tabRooms` instead of re-seeking.
4. **Color tuning**: red `&HFF&` / white `&HFFFFFF` live in the recolor block in both forms'
   `MapMapExits`. Known+up/down rooms lose the up-vs-down color distinction by design (tooltip
   still lists exits).
5. **Hash false positives** are inherent (8-char fingerprint, not unique) — same limitation as
   the existing Find button; tooltip name lets the user verify.

## Suggested commit

Once built/verified, commit the 6 files together (this is a source change, not a release, so
`mudexplr.exe` need not be rebuilt unless doing a release). Example message:
`Add "Mark MegaMUD Rooms" map option (recolor rooms found in rooms.md)`
