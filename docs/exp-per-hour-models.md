# Exp/Hour Prediction Models (`modExpPerHour.bas`)

`CalcExpPerHour()` estimates effective experience/hour for a lair-style farming
zone. It runs one or more **selectable models** and averages the enabled ones.
Output is a `tExpPerHourInfo` (exp/hr plus attack / move / HP-rest / mana-rest /
roam / overkill / slowdown fractions) consumed by the monster list and tooltips.

## The models

| Model | Approach | Status |
|---|---|---|
| A (`ceph_ModelA`) | Closed-form: q-elasticity HP recovery, mana-pool model, density-based movement, overlap credits. | Legacy. Accurate but internally tangled. |
| B (`ceph_ModelB`) | Smoothed, band-aware closed form. **Dozens of hand-tuned SmoothStep/BandWeight constants fit to specific zones.** | Legacy. **Overfit - do not extend.** |
| C (`ceph_ModelC`) | Per-lair "cycle" macro-simulation (combat profile -> cycle profile -> per-hour). | Legacy; basis for D. |
| D (`ceph_ModelD`) | **Round-by-round simulation** in real game ticks. Uses the **canonical `nRTK`** (from `CalcCombatRounds` via `GetLairInfo`, same input A/B use), then simulates HP/MP drain + recovery + movement. | **Recommended.** Opt-in, off by default. |

Model flags: `bGlobal_cephModelA/B/C/D`. Each maps to `chkEPH_Model(0/1/2/3)` in
frmSettings and INI key `cephModel{A..D}`.

## Model D design (the recommended one)
- **Combat rounds come from the canonical `CalcCombatRounds`** (`modMMudFunc.bas`),
  passed in as `nRTK` (rounds-to-kill a *single* mob, with the 0.5-round rule,
  surprise credit and engine-correct mob-HP-regen already applied by `GetLairInfo`).
  `RTC = nRTK * nNumMobs`. This is the same RTK Models A/B consume. If `nRTK` is 0
  (e.g. some harness rows), D derives it with the same 0.5-round rule Model A uses.
  D **no longer** calls `cephC_BuildCombatProfile` — that function is now Model-C-only.
- Round-by-round incoming damage with true multi-mob ramp-down and a **per-round**
  damage threshold (a heal that covers one mob but not the whole pack is modeled
  correctly).
- **Serialized recovery**: rest (HP) XOR meditate (MP), never together, never in
  combat; passive HP/MP always tick.
- Movement = measured `nAvgWalk * nWalkSpeed` (no density heuristics).
- Only tunables: the four user knobs `nGlobal_ceph{XP,DMG,Mana,Move}_Knob` plus
  three named constants in the declarations block:
  - `cephD_KILL_OVERHEAD_SEC` (1.5) - per-kill looting/retarget/latency; gated to
    fade when recovery downtime already absorbs it.
  - `cephD_HEAVY_REST_RELIEF` (0.35) - players fight in a lower HP band on brutal
    fights instead of topping to the rest target each lair.
  - `cephD_MEDITATE_EFF` (0.5) - meditate skill efficiency (interrupts); passive
    MP is unaffected.

## Confirmed game mechanics (used by the formulas; not obvious from code)
- **`nCharHPRegen` arrives in _resting-rate_ form (already game-value x3).** So
  passive HP = `HPRegen/3` per 30s; rest HP = `HPRegen` per 20s. Mana
  (`nCharMPRegen`, `nMeditateRate`) is passed as-is.
- **Recovery is serialized** - rest (HP) and meditate (MP) are mutually exclusive
  and never happen in combat. Passive HP/MP tick at all times.
- **`nMobDmg` (= `nAvgDmgLair`) already bakes in multi-mob ramp-down** - in
  `GetLairInfo` (`modMMudDatabase.bas`) per-mob `AvgDmg` is multiplied by `nRTK`
  then divided by `avgAlive = (nMaxRegen+1)/(2*nMaxRegen)`. Model D reverses this
  to recover clean per-mob/round damage.
- Tick constants: round 5s, rest 20s, regen 30s, meditate 10s. Mob HP-regen
  cadence: `STOCK_MOB_HPREGEN_ROUNDS=18`, `GMUD_MOB_HPREGEN_ROUNDS=6`
  (`modMMudFunc.bas`), gated by `bGreaterMUD`.
- Inputs come from `GetLairInfo`/`LairInfoType` + `tCharacterProfile`; per-argument
  meaning is documented in the `CalcExpPerHour` header comment.

## Calibration harness
`RunAllSimulations` (in `modExpPerHour.bas`) runs ~18 real in-game observations
embedded in its `SIM_TABLE` and prints `Avg Exp/Rest/Mana/Move Diff`.
- To test one model in isolation, enable **only** its checkbox; turn on
  `bDebugExpPerHour` (and `DEVELOPMENT_MODE`) for per-sim detail and a `D:`-style
  per-model column with "Show in Detail".
- **GOTCHA: the harness "Avg Diff" is signed (`1 - obs/est`) = bias, not accuracy.**
  A near-zero average can hide large offsetting errors - compute mean-absolute
  error (MAE) separately when comparing models.
- Reference standings on the 18 rows (all knobs = 1.0): A ~11.2% MAE, D ~11.6%,
  B ~13.8%, C ~16.5%. (Signed bias: A -2.1%, D +3.0%, B +12.6%, C +13.3%.)

## Adding a model (the wiring checklist)
A new model touches six places:
1. `modExpPerHour.bas`: `Global bGlobal_cephModel<X>`; the `ceph_Model<X>`
   function; its call + a `Case n` arm in the averaging/show-all loop in
   `CalcExpPerHour`; and `DebugPrintExpHrGlobals`.
2. `frmSettings.frm`: `chkEPH_Model(n)` in cmdDefaults / PopulateExpFields /
   cmdSave_Click (including the "all off" fallback) / the WriteINI block.
3. `frmMain.frm`: the startup default; the INI read (default the key so existing
   `settings.ini` files stay compatible); and the `bSettingsPass` "defaults" gate.

## Editing reminder
`modExpPerHour.bas` and the `.frm` files are Windows-1252 + CRLF with high-byte
characters in comments. Edit them via the Latin-1 (cp28591) PowerShell round-trip
described in the repo `CLAUDE.md` - never the UTF-8 Edit/Write tools.

## Audit outcome (2026-06-20) - D decoupled from C's combat core
The audit compared `cephC_BuildCombatProfile` line-by-line against the canonical
`CalcCombatRounds` (`modMMudFunc.bas`), the engine the rest of the app uses. Core
finding: there were **two divergent rounds-to-kill engines** - A/B consumed the
canonical `nRTK`; C/D recomputed their own in `cephC_BuildCombatProfile`, which had
drifted. Verified discrepancies in the cephC recompute:
1. **One-shot overkill bug** - one-shot branch set `hpBeforeLast = 0` → 100% overkill
   on every one-shot (should be `mobHP`). Display-only.
2. **Min-damage tail applied unconditionally** (`extraProbNormal`) - added to RTK
   whenever `minDmg < avgDmg` with no round-boundary gate; canonical gates strictly.
   Feeds exp/hr.
3. **Integer-ceil RTK** vs canonical round-up-to-0.5 - systematic upward RTK bias
   (e.g. SIM5 perMob 292 HP / 232 dmg: cephC 2.0 vs canonical 1.5).
4. **Mob-HP-regen gate hardcoded `RTK>=6`** vs canonical `0.9 x regenRounds`
   (16.2 stock / 5.4 GMUD) - over-applies regen on stock long fights.
5. **Surprise-miss = full wasted round** vs canonical scaled credit - likely too harsh.

Calibration blind spots that hid #2/#4/#5: all 18 SIM rows have `MinRoundDMG==CharDMG`
(no variance), none trigger mob-regen, only SIM18 has surprise (at 100%).

**Resolution (implemented):** Model D now uses the canonical `nRTK` it already
receives (`RTC = nRTK*nNumMobs`), keeping its own HP/MP-drain + serialized-recovery
+ movement loop. It no longer calls `cephC_BuildCombatProfile`. Overkill (display) is
computed by a small `cephD_OverkillFrac` helper with the one-shot bug fixed; slowdown
is `(nRTK-1)/nRTK`. This collapses the two-source problem and removes #2-#5 from D's
exp/hr path in one move.

**Still open / Model C only:** `cephC_BuildCombatProfile` is now used **only by
Model C** and still contains bugs #1-#5. Since C is legacy, these were left in place;
fix or retire C separately. Re-run `RunAllSimulations` (D-only checkbox) after the
next IDE build to re-measure D's MAE - the change shifts several rows (SIM1 1.0→1.5,
SIM5/SIM17 2.0→1.5), so the prior ~11.6% figure is stale.
