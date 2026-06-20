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
| D (`ceph_ModelD`) | **Round-by-round simulation** in real game ticks. Reuses `cephC_BuildCombatProfile` for RTK/RTC/overkill, then simulates HP/MP drain + recovery + movement. | **Recommended.** Opt-in, off by default. |

Model flags: `bGlobal_cephModelA/B/C/D`. Each maps to `chkEPH_Model(0/1/2/3)` in
frmSettings and INI key `cephModel{A..D}`.

## Model D design (the recommended one)
- Full combat data: backstab chance/min/max, first-round, min-round damage,
  engine-gated mob HP regen.
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
