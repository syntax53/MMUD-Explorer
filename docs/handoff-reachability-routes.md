# Handoff: monster reachability + route feature

Status as of 2026-10-12 (session 8509831f). Branch: `@syntax53/261005/pr-test/justin-mob-reachability`
(a test of PR syntax53/MMUD-Explorer#187, "Add monster filtering based on level gating").
This file is untracked scratch for continuity; delete or commit at the end.

## Committed so far (on top of the PR merge)
- `6fc8627` Show monster reachability in the detail pane — "Reachable | Yes/No, blocked by ..." rows,
  `ReachMonsterInfo` + gate tracer in `modReachability.bas`, `reach_compute` double-click.
- `70c0caa` Show a route to each monster in the detail pane — `ReachRoute`, edge kinds/labels,
  seeded BFS with parents, points of interest (Godfrey, Arlysia 17/2435, Khazarad 6/1334,
  Lost City 16/426, Albion 14/1, Port Blackwater 14/10), waypoint rows every N (Settings option
  `RouteWaypointInterval`, dialog grown), results window via double-click on Yes/No.

## Round 4 — adversarial audit fixes (2026-10-12, session f70ac558)
STATUS: APPLIED, UNCOMPILED (no VB6 on this PC). Three Opus reviewers audited the whole branch; fixes:
- modReachability: level **breakpoints now include random-group edges** (they were skipped, so a band
  could be evaluated at a level where a level-gated group was unusable → filter/detail pane disagreed
  with the route finder); `maxlevel 0` in a textblock = no upper limit (as for exits); item-spell edges
  carry the chain's `price`; `teleport`/`summon` keyword match is case-insensitive;
  `ReachEnsureComputed` has a re-entrancy guard (`bReachComputing` — its DoEvents let the filter form's
  Apply start a nested build), treats any build step's error as failure (`bBuildError`, set in each
  step's handler; previously partial data was marked computed), and restores frmMain.Enabled to its
  previous value; `VisitAtLevel` caches its last result per level (the detail pane needs it 3× per
  monster); **`ReachMonsterInfo` gained an 11th arg `nReachableAtLvl`** (lowest higher band that reaches
  the monster, 0 = none) and `bRandomGroup` is now True only for an *unusable* group — a level-gated
  usable group reports the group's combined level range.
- modMain: reach block extracted from `PullMonsterDetail` into `AddMonsterReachRows` (64K
  compiled-code-per-procedure risk); the No row's `reach_route` tag now follows the FULL route (what
  double-click opens), not the to-gate route; the No text appends "(reachable from Level X)" when a
  further gate raises the real requirement, or "(still blocked at higher levels)"; `RouteLabelNames`
  caps digit runs at 9.
- frmMain: `reach_compute` rows call `ReachRecompute` (a failed build was never retried) + message;
  pre-existing compare-pane bug (greet_text/textblock read `lvMonsterDetail`) fixed; Calculate
  Reachability menu skips the main-list refresh when the reachable-only filter is on (FilterMonsters
  refills the list and clicks a row itself; the compare pane still refreshes); results caption gets a
  ToolTipText (one-line label clips) and singular "step"; interval default 0 → 25 (was clamped to 1).
Known, deliberately not changed: the detail pane's waypoints go to the nearest gate (reverse-BFS) while
the No-row results window shows the cheapest ignore-level route, which may cross a different gate;
room-spell random blocks with several `random` ops merge into one group (LineNo = 0).

## Round 3 — coin-detour cap + alternate-route row (2026-10-12, session f70ac558)
STATUS: COMMITTED as `6194f3c`, UNCOMPILED. **VB6 is not installed on the user's current PC**, so
nothing was compiled or run. Only static checks were done: high-byte counts unchanged (0/1/6), zero lone
LFs, and stripped diffs show only the intended hunks. The user compiles and tests in the IDE.
Files: modReachability.bas, modMain.bas, frmMain.frm, docs/changelog.txt, README.md.
- `RouteSearch` gains `bAvoidCoin` + `cost()` output. `ReachRoute` runs it twice: F (fewest coin steps,
  then steps) and S (fewest steps).
  - It uses S when `nCoinS > nCoinF And nS < nF And nF > ROUTE_DETOUR_MAX(3) * nS`, otherwise F.
  - The other route is the "alternate".
  - New signature (17 args): `bAltRoute` after `bIgnoreLevel`, then `nCoinSteps, nAltSteps, nAltCoinSteps,
    sAltSourceName` at the end. `bAltRoute=True` returns the alternate (False if none).
  - New helpers: `RouteDest` (category → cost → index, so a coin-free destination beats a closer paid one),
    `RoutePathInfo`, `RouteSteps`, `POIName`.
- modMain:
  - `RouteCoinText` ("1 toll/fee"), `RouteKindText` ("Coin-free route" / "Shorter route" /
    "Route with fewer tolls/fees"), and private `AddRouteAltRow` (Tag `reach_route_alt`, gray).
  - Yes row appends ", N toll(s)/fee(s)". The Yes alt row goes after the route rows.
  - The No case makes an extra `ReachRoute(bIgnoreLevel=True)` call. Its alt row reads
    "... past the level gate: N steps from X (double-click to view)".
- frmMain: `ShowMonsterRoute(nMon, Optional bAltRoute)`. The caption names the route kind, adds the coin
  count, and mentions the alternate. `reach_route_alt` cases are added in both detail double-click
  handlers.
- Expected for 2352 at L49:
  - No row unchanged.
  - Alt row "Coin-free route past the level gate: 187 steps from Bank of Khazarad".
  - No-row results: the 30-step ship route from Port Blackwater.
  - Alt-row results: the 187-step route.

## Round 2 — the user's eight requests
STATUS: ROUND 2 COMMITTED as `3ecfcf8` (after the user's own `fb4613c` "Enable dev mode and ignore
reachability CSV"). The coin-cap question it left open is answered by round 3.
(previous status kept below)
STATUS 2026-10-12 04:20: ROUND 2 COMPLETE AND VERIFIED IN-APP (second compile after the two
follow-up edits succeeded; captain 2352 No-row results now run from Bank of Khazarad with the full
route, Step column first, "(Level 40+)" labels, auto rows every 25 silent steps). ONLY REMAINING STEP:
commit the 7 changed files (README.md, docs/changelog.txt, frmMain.frm, frmResults.frm,
frmSettings.frm, modMain.bas, modReachability.bas) once the user says so. Observation to raise with
the user: coin avoidance has no length cap, so the ignore-level route to Kingsport is a 187-step
coin-free detour through the Negative Power Plane portals instead of the 30-step 2,000,000-gold ship.
PROGRESS (earlier): ALL agent work landed; first scratch compile succeeded; in-app checks passed for
the menu item, wolfman (lair destination 17/2613), spectral assassin / doom chicken (no bogus rooms),
captain 2352 No row. Two Fable follow-up edits after that: (a) modReachability `ReachRoute` now picks
sources at the character's level even when bIgnoreLevel (otherwise the No-row results started from
Bank of Albion beyond the gate and hid the ship step); (b) frmResults Form_Resize stretches
lblCaption. Remaining: recompile, re-check the captain's No results window (expect route from Port
Blackwater through the `secure passage to albion` step with "(Level 50+)" and "(price 2,000,000)"),
then commit everything (6 source/doc files + README). Earlier history below.
Agent A (modReachability.bas) FINISHED and reviewed (encoding checks pass).
Agent B was cut off by the usage limit after Part A: modMain.bas now has `FormatRouteLabel`
(+ private helpers `RouteLabelNames`, `RouteLabelClassRace`), the new `ReachRoute` signature in
`PullMonsterDetail`, and the single-row No layout. A second agent was launched for Parts B-E
(frmMain.frm ShowMonsterRoute two-column rows + new signature + default 25 + menu item renumbering;
frmSettings tooltip; frmResults Form_Resize last-column sizing; changelog/README). If that agent
also died, `git diff --stat` tells which of those files changed; frmMain.frm MUST be updated to the
new 12-arg `ReachRoute` signature or the project will not compile.
Full design: `C:\Users\matthew_snead\.claude\plans\this-branch-is-a-virtual-whisper.md`.
Two Opus agents were launched in parallel (no shared files):
- **Agent A — `modReachability.bas`**: (1) skip references followed by a `+` truncation marker in
  `NextMapRoom`/`ParseRefs` (fixes spectral assassin 2658 → 14/62 and doom chicken 2650 → 14/59,
  whose spells 5623/5615 have a cut-off `Casted By` ending `Room 14/62+`); (2) `mRooms` values =
  category (1 Room, 2 Group(lair), 3 spell/textblock-derived incl. death summons, 4 Group) and the
  destination picks lowest category, then distance (fixes wolfman 807 choosing 17/2289 via monster
  793's death-summon spell over its 46 lairs); (3) `eCoin()` per edge from `(Toll: N)` exits and
  `price N` textblock ops, `REACH_STEP_COIN = 32`, label suffix `(price 2,000,000)`;
  (4) `RouteSearch` = Dijkstra keyed `coinSteps*1e6 + steps` (coin-free paths preferred);
  (5) `ReachRoute` gains `ByVal bIgnoreLevel As Boolean` after `bToGate` (full route past level
  gates for the No row's results window); (6) labels carry raw numbers (`spell #N`, `item #N`,
  raw exit restriction text) for the UI to name; (7) `Public Sub ReachRecompute()`.
- **Agent B — UI**: `modMain.bas` `FormatRouteLabel` (names via GetItemName/GetSpellName/
  GetClassName/GetRaceName per `bHideRecordNumbers`; room numbers always), detail pane No row
  becomes one row `No - blocked by <gate> (m/r), needs Level 50+` (blocked-by and leads-to rows
  removed; waypoints to the gate kept); `frmMain.frm` `ShowMonsterRoute` two-column results
  ("Step" | "Location"), auto rows only after N silent steps, no "(waypoint)" text, No case uses
  `bIgnoreLevel=True`; default interval 25 (`Form_Load`) + frmSettings tooltip; new Options menu
  item "Calculate &Reachability" at `mnuOptionsItems` Index 1 with Index 1..12 renumbered to
  2..13 (designer block + `Select Case` + refs `mnuOptionsItems(2).Enabled`×3 → (3),
  `mnuOptionsItems_Click(2)`→(3), `(1)`→(2)); `frmResults.frm` `Form_Resize` sizes the LAST
  column; changelog NEW/UP/FIX lines + `docs\sync-readme.ps1`.

If this session died mid-way: check `git status`/`git diff` to see which of the above landed, read
the plan file, and finish the rest. Agents' scratch .ps1 scripts live in the session scratchpad
(may be gone); re-apply edits with the Latin-1 method.

## After the agents: review + verify recipe
1. Encoding checks on every touched .bas/.frm: high-byte count equals HEAD
   (`git show HEAD:f | tr -cd '\200-\377' | wc -c` vs `tr -cd '\200-\377' < f | wc -c`), zero lone LFs
   (`python -c "d=open('f','rb').read();print(d.count(b'\n')-d.count(b'\r\n'))"`), stripped diff
   (`git diff -- f | tr -cd '\11\12\40-\176'`) only in intended hunks.
2. Compile to scratch (never the committed exe). VB6.EXE has a RUNASADMIN compat flag, so in
   PowerShell: `$env:__COMPAT_LAYER='RunAsInvoker'; Start-Process 'C:\Program Files (x86)\Microsoft Visual Studio\VB98\VB6.EXE' -ArgumentList '/make','"<repo>\MMUD_Explorer.vbp"','/outdir','"<scratch>"','/out','"<scratch>\build.log"' -Wait`
   → log says "Build of 'mudexplr.exe' succeeded.". Keep `Remove-Item` out of the same command
   (tool safety filter). Afterwards `git checkout -- MMUD_Explorer.vbp MMUD_Explorer.vbw` if the
   compile rewrote them, and delete any stray `data-*.ldb`.
3. Drive the exe with the Paradigm DB (`I:\My Drive\mud\dats\data-Paradigm-1.9.1.mdb`), see memory
   `vb6-app-ui-driving` (mscomctl lists need real mouse clicks + PrintWindow screenshots; UIA
   assembly makes the shell DPI-aware → mixed coordinate modes). Acceptance:
   - 2658 spectral assassin / 2650 doom chicken: no 14/62 / 14/59 destination.
   - 807 wolfman: destination is a lair room.
   - 2352 old sea captain, level 49: `No - blocked by Blackwater Harbor, Wharf (14/759), needs Level 50+`;
     double-click → results through to Kingsport Harbor, Wharf (14/702) with the
     `secure passage to albion` row showing `(Level 50+)` and `(price 2,000,000)`. Level 50 → `Yes`.
   - Results: Step column first, no "(waypoint)", names for items/keys/spells, auto rows only after
     25 silent steps. Options menu has "Calculate Reachability" above "Calc. Monster Dmg vs Char";
     Settings/Reload/Remove Filters still work. Settings tooltip says Default = 25.
4. Commit all touched files together (changelog + README in the same commit), message ends with
   `Co-Authored-By: Claude Fable 5.1 <noreply@anthropic.com>`. Do NOT commit mudexplr.exe
   (release builds are the user's IDE build with the four dev flags off).

## Decisions already made with the user (do not re-ask)
- Yes/No verdict and "blocked by" gate stay Godfrey-based; POIs only affect routes.
- Unreachable: single No row with blocked-by text; detail waypoints stop at the gate; results
  window goes all the way through as if no level limit.
- Destination priority Room > Group(lair) > spell/textblock-derived (incl. death summons) > Group.
- No level-at-compute flag/prompt (every level band is precomputed); the menu item stays for manual
  recompute and for future class/race/alignment gates (then: record a char config string at compute
  time and prompt in `FilterMonsters` like `sMonsterDamageVsCharDefenseConfig`).
- Coin avoidance = toll exits + priced commands only (not item-consuming steps).
- Coin avoidance is capped: if the coin-free route is more than 3× the shorter coin route, show the coin
  route. Either way, offer the other route as a detail row (double-click opens it in the results window).
  Coin is shown as a count of paid steps, not a gold total, because toll and `price N` units may differ.
  The No case's alt row covers the full route past the gate (what the No row opens). The results window
  has no switch row of its own: `SetupResultsWindow` unloads the form, so it can't reload from its own
  click handler.
- Interval default 25; detail pane caps waypoint rows at 10 then "... plus N more".

## Known facts worth keeping
- CLAUDE.md file rules: .bas/.frm are Windows-1252 + CRLF; never Edit/Write them; Latin-1 round-trip.
- `GotoLocation` (frmMain) needs results rows whose text starts with "Room" and Tag "m/r";
  with subitems it reads `ListSubItems(1)`.
- VB6 menu control arrays must be contiguous ascending in designer order.
- Open PR #187 is RJCo's; the user (syntax53) is testing it on this branch.
