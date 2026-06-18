# Highlight Known MegaMUD Rooms on the Map

## Context

MMUD Explorer draws area maps into `picMap` (frmMain) and `picMap`/`picZoomMap` (frmMap),
overlaying shapes for room features (commands, NPCs, lairs, shops, spells, exits). Players use
the **MegaMUD** client, whose `rooms.md` files list every room the community has charted, keyed
by the same room-hash this app already computes (`Get_MegaMUD_RoomHash` + `Get_MegaMUD_ExitsCode`
in `modMain.bas`). Today the only consumer of that hash is the `cmdMapMegaRoomFind` button, which
scans `rooms.md`/`.mp` files to locate *one* current room.

This feature lets the user see, at a glance, **which rooms on a drawn map are already known to
MegaMUD**. We parse all `rooms.md` files once into an in-memory database, then while drawing a
map we highlight any cell whose computed hash is present in that database.

We repurpose the seldom-used `chkMapOptions(12)` ("Don't Follow Restricted") checkbox to toggle
this, dropping the old restricted-exit behavior (defaulting it off).

Scope: **both** frmMain and frmMap (incl. the zoom map). Data is held **in memory per session**.
Matched cells get a **gap-frame highlight** and **tooltip info**.

## Key existing code (reused, do not rewrite)

- `Get_MegaMUD_RoomHash("", map, room)` (3 hex chars) `& Get_MegaMUD_ExitsCode(map, room)`
  (5 hex chars) = the 8-char value matching field `sArr(0)` in `rooms.md`. This exact
  concatenation is already used in `cmdMapMegaRoomFind_Click` (`frmMain.frm` ~19762).
- `rooms.md` line format (split on `:`): `sArr(0)`=hash, `sArr(5)`=code, `sArr(6)`=group,
  `sArr(7)`=name. See `MapMegaScanRoomsDatabase` (`frmMain.frm:19503`) for the parse pattern.
- `MapMapExits` sets each cell's up/down background color at `frmMain.frm:33639-33647`
  (mirror in `frmMap.frm`) — the insertion point for the highlight + tooltip.
- `MapDrawOnRoom` shape renderer: `frmMain.frm:32752`, `frmMap.frm` ~45228 (zoom-scaled).
  `EnumDrawRoom` enum (`frmMain.frm:19141`, `frmMap.frm:43638`) currently 0–13.
- File picking: `frmMain.oComDag` CommonDialog control (used by `cmdMapMegaRoomFind`).
  `FileSystemObject` recursion pattern also in that routine. INI helpers `ReadINI`/`WriteINI`
  in `modSettings.bas`.

> **Editing rule:** `.frm`/`.bas` are CP1252. Do NOT use Edit/Write (strips high bytes). Use the
> Latin-1 (cp 28591) PowerShell round-trip per CLAUDE.md / `[[edit-tool-strips-high-bytes]]`, then
> verify high-byte count vs HEAD.

## Implementation

### 1. Global in-memory database (`modMMudDatabase.bas`, near other globals)
```vb
Public dictMegaRooms As Scripting.Dictionary   'key=8-char hash, value="code|group|name"
Public sMegaRoomsRootFile As String            'remembered root rooms.md path (session only)
```

### 2. Shared build routine (`modMain.bas`, alongside the hash functions)
- `Public Function MegaRooms_EnsurePopulated(ByVal bForceRepick As Boolean) As Boolean`
  - If `dictMegaRooms` already has entries and not `bForceRepick`: prompt
    "Refresh known-rooms data from rooms.md?" (Yes = re-pick + rescan, No = keep). 
  - If empty (or refresh requested): MsgBox guidance (reuse the wording from
    `cmdMapMegaRoomFind`), then open `frmMain.oComDag` filtered to `Rooms.md`
    (init dir from `ReadINI("Settings","Last_MegaMUD_DBFolder")`, same fallbacks as the
    existing button). On cancel, return False.
  - Call `MegaRooms_ScanFolderRecursive` on the chosen file's parent folder, store the root
    path in `sMegaRoomsRootFile`, save folder to `Last_MegaMUD_DBFolder`. Return True.
- `Private Sub MegaRooms_ScanFolderRecursive(oFolder As Folder)` — for each file named
  `rooms.md` (case-insensitive) call `MegaRooms_ParseFile`; recurse into every subfolder
  (full depth, unlike the button's 2-level limit).
- `Private Sub MegaRooms_ParseFile(ByVal sPath As String)` — mirror
  `MapMegaScanRoomsDatabase`'s split, but for every line with `UBound(sArr) >= 7` do
  `If Not dictMegaRooms.Exists(sArr(0)) Then dictMegaRooms.Add sArr(0), sArr(5) & "|" & sArr(6) & "|" & sArr(7)`
  (keep first on hash collision — fine for existence highlighting).
- Initialize `Set dictMegaRooms = New Scripting.Dictionary : dictMegaRooms.CompareMode = vbTextCompare`
  before first fill.

### 3. New draw shape — gap-frame (both forms)
- Add `drMegaKnown = 14` to both `EnumDrawRoom` enums.
- Add `Case 14` to both `MapDrawOnRoom` subs: a box **outline** (`B` flag) drawn in the gap,
  one gap-half outside each cell edge:
```vb
Case 14: 'mega known room
    picMap.DrawWidth = nSize
    x1 = oLabel.Left - nMapCellGapDraw
    y1 = oLabel.Top - nMapCellGapDraw
    x2 = oLabel.Left + oLabel.Width + nMapCellGapDraw + nMapCellGapDrawAdj
    y2 = oLabel.Top + oLabel.Height + nMapCellGapDraw + nMapCellGapDrawAdj
    picMap.Line (x1, y1)-(x2, y2), QBColor(nColor), B
```
  In `frmMap`'s version, mirror the zoom handling used by other cases (`Set oPM`, `nAltSize`).
  Pick a distinct unused color (e.g. bright white/orange) finalized against the cell backgrounds
  during testing.

### 4. Highlight hook in `MapMapExits` (both forms, just after the up/down color block
`frmMain.frm:33647`)
```vb
Dim sMegaInfo As String   'declare with the other locals
...
If chkMapOptions(12).Value = 1 And Not dictMegaRooms Is Nothing Then
    If dictMegaRooms.Count > 0 Then
        Dim sMegaHash As String
        sMegaHash = Get_MegaMUD_RoomHash("", Map, Room) & Get_MegaMUD_ExitsCode(Map, Room)
        If dictMegaRooms.Exists(sMegaHash) Then
            Call MapDrawOnRoom(lblRoomCell(Cell), drMegaKnown, 3, <color>)
            sMegaInfo = dictMegaRooms(sMegaHash)   'code|group|name
        End If
    End If
End If
```
- In the tooltip block (`If chkMapOptions(5).Value = 0`, ~33653) append when `sMegaInfo <> ""`:
  parse `code|group|name` and add a line like
  `"MegaMUD: [" & group & "] " & code & " - " & name`.
- The hash funcs re-`Seek` `tabRooms`; this runs after the color block and the tooltip block
  reads from already-captured locals, so position is safe.

### 5. Repurpose the checkbox
- **Captions** → e.g. `"Mark MegaMUD Rooms"`: `frmMain.frm:7628`, `frmMap.frm:10442` (designer
  Caption lines — edit text only, leave the control block structure intact).
- **Drop "Don't Follow Restricted":** remove the restricted-skip lines `frmMain.frm:33611` and
  `frmMap.frm:44926` (restricted exit types 13–15 now always followed = old default-off behavior).
  Remove frmMap's persistence: load `frmMap.frm:43782` and save `frmMap.frm:45798`
  (`ExMapNoRestricted`). Index 12 is **not** persisted on either form, so it starts unchecked
  each session — matching the in-memory-per-session model (enabling it triggers the file
  prompt/scan).
- **Click handlers:**
  - frmMain has **no** `chkMapOptions_Click`; add `Private Sub chkMapOptions_Click(Index As Integer)`
    handling `Index = 12`: when checked, `If MegaRooms_EnsurePopulated(False) Then` redraw if a map
    is up (`If nMapStartMap > 0 Then MapStartMapping nMapStartMap, nMapStartRoom`); if the user
    cancels the picker, reset the checkbox to 0.
  - frmMap already has `chkMapOptions_Click` (`frmMap.frm:43800`, handles 6/7) — add an
    `Index = 12` branch doing the same ensure-populate + redraw (its redraw entry, e.g.
    `ResizeMap`/`MapStartMapping`).

## Files touched
- `modMMudDatabase.bas` — globals (dict, root path).
- `modMain.bas` — `MegaRooms_EnsurePopulated`, `MegaRooms_ScanFolderRecursive`,
  `MegaRooms_ParseFile`.
- `frmMain.frm` — enum member, `MapDrawOnRoom` Case 14, `MapMapExits` highlight+tooltip,
  remove line 33611, caption 7628, new `chkMapOptions_Click`.
- `frmMap.frm` — enum member, `MapDrawOnRoom` Case 14 (zoom), `MapMapExits` highlight+tooltip,
  remove lines 44926/43782/45798, caption 10442, `chkMapOptions_Click` Index 12 branch.
- `docs/changelog.txt` + `README.md` top — user-facing note.

## Verification
1. Build `mudexplr.exe` in the VB6 IDE (open `MMUD_Explorer.vbp` → Make). No CLI build/tests exist.
2. Load a known `.mdb`, draw a map for an area present in the sample `rooms.md`.
3. Check "Mark MegaMUD Rooms" → confirm the file prompt appears, select the root `rooms.md`,
   confirm the map redraws with gap-frames around rooms that exist in `rooms.md` and not around
   rooms that don't. Hover a framed room → tooltip shows the MegaMUD code/group/name.
4. Toggle off → frames disappear on redraw. Toggle on again → "refresh?" prompt (No keeps data).
5. Repeat on frmMap at a normal zoom and a `picZoomMap` zoom level; confirm the frame scales.
6. Confirm restricted exits (class/race/level, types 13–15) are now always followed on both forms.
7. Post-edit: verify CP1252 integrity — high-byte count unchanged vs HEAD, no lone LFs, and a
   high-byte-stripped diff shows only intended changes.
