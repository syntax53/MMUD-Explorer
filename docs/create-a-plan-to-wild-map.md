# Automated DupeExpRequired Reduction (DupeExpReduction / DupeExpMinimum)

## Context

`DupeExpRequired` (INI, long, default 8,000,000,000) gates creation of additional character slots. The gating is **per-slot and sequential**: [GMUDCharacterSelectionMenu.cs:87-106](GreaterMUD.Module/Menu/GMUDCharacterSelectionMenu.cs) computes XP needed for slot N as `DupeExpRequired - (XP on slot N-1)`. Reaching the threshold on character 1 only unlocks slot 2; slot 3 requires the threshold on character 2, etc. `GetTotalSlotXp` (lines 120-144) uses the max XP **ever** achieved on a slot (from `StatePlayers`, which includes rerolled history), so unlocked slots survive rerolls.

We want the realm to automatically lower the effective requirement as more characters achieve it:

```
effective = max(DupeExpMinimum, DupeExpRequired - qualifyingCount * DupeExpReduction)
```

(The request wrote `min()`, but "down to a minimum of DupeExpMinimum" means a clamp — `max()`.)

**Confirmed decisions (user):**
- `qualifyingCount` compares each character's XP against the **base INI** `DupeExpRequired` (no feedback loop from the reduced value).
- Count **each** active, non-`Rerolled` character (one account with 3 qualifying chars counts 3). `FromReroll` chars are active and count if they qualify.
- Compute **live, on demand** — no timer/cache. The player list is in-memory; counting is cheap.
- With `DupeExpReduction = 0` (unset), behavior must be byte-identical to today.

## Changes

### 1. [GMUDServerSettings.cs](GreaterMUD.Module/GMUDServerSettings.cs) — two new settings

- Properties next to `DupeExpRequired` (~line 637):
  ```csharp
  public long DupeExpReduction { get; private set; } = 0;
  public long DupeExpMinimum { get; private set; } = 0;
  ```
- Two new `case` blocks in `CheckSetting()` after the `DupeExpRequired` case (lines 463-471), copying its exact pattern (`long.TryParse` into existing `tempLong`, `>= 0 && <= long.MaxValue` range check).

### 2. [GMUDServer.cs](GreaterMUD.Module/GMUDServer.cs) — fields, constructor copy, effective-value method

- After `public long DupeExpRequired;` (line 179): add `public long DupeExpReduction;` and `public long DupeExpMinimum;`
- After line 271 in the constructor: copy both from `inSettings`.
- New method near `GetTopPlayersList` (line 1301):
  ```csharp
  public long GetEffectiveDupeExpRequired()
  {
      // Disabled requirement or no reduction configured: behave exactly as before.
      if (this.DupeExpRequired <= 0 || this.DupeExpReduction <= 0)
          return this.DupeExpRequired;

      long qualifyingCount;
      lock (this.AllPlayers)
      {
          qualifyingCount = this.AllPlayers.Values.Count(x => !x.Rerolled && x.Experience >= this.DupeExpRequired);
      }

      long minimum = this.DupeExpMinimum > 0 ? this.DupeExpMinimum : 0;

      // Clamp before multiplying so a huge reduction value can't overflow.
      if (qualifyingCount >= (this.DupeExpRequired / this.DupeExpReduction) + 1)
          return minimum;

      return Math.Max(minimum, this.DupeExpRequired - (qualifyingCount * this.DupeExpReduction));
  }
  ```
  Notes:
  - `DupeExpRequired == 0` returns 0 immediately — menus treat 0 as "no requirement", and `DupeExpMinimum` must never resurrect a disabled gate.
  - The divide-and-compare guard short-circuits to the minimum before `count * reduction` could overflow.
  - `System.Linq` is already imported in GMUDServer.cs; `lock (this.AllPlayers)` matches the convention used by `GetTopPlayersList`.

### 3. [GMUDCharacterSelectionMenu.cs](GreaterMUD.Module/Menu/GMUDCharacterSelectionMenu.cs) — use the effective value

Compute once per user action and thread it through as a parameter (avoids re-counting per slot in the render loop):

- `GetXpNeededForSlot(List<Player> myCharacters, int slot)` → add `long dupeExpRequired` parameter; replace `this.GMUDServer.DupeExpRequired` at lines 92 and 99 with it.
- `GetTotalSlotXp(...)` → same parameter; use it in the `max >= ...` comparison at line 130 (consistent threshold for "did this slot ever qualify").
- `ShowCharacters` (line 20): `long dupeExpRequired = this.GMUDServer.GetEffectiveDupeExpRequired();` once before the slot loop; pass into the `GetXpNeededForSlot` call at line 44.
- `Process` slot-selection branch (line 169): compute fresh and pass in.
- **Gate creation only, never usage of existing characters** (confirmed requirement): the effective requirement can *rise* when qualifying characters are deleted/rerolled, and a player must never be locked out of a character they already created. In `Process`, restructure the line 169 block: move the `selectedPlayer` lookup (currently line 171) *before* the XP check, and apply the gate only when the slot is empty:
  ```csharp
  Player selectedPlayer = myCharacters.Where(x => x.CharacterSlot == slotSelected).FirstOrDefault();
  if (selectedPlayer != null || GetXpNeededForSlot(myCharacters, slotSelected, dupeExpRequired) <= 0)
  ```
  (This also fixes the same latent lockout in current code when an admin raises `DupeExpRequired` in the INI.) `ShowCharacters` needs no change for this — the "Exp needed" line is already only rendered for empty slots (line 63 is inside the `character == null` branch).

**Note on sequential gating:** unchanged by this plan. Slot N is still gated by XP on slot N−1; only the threshold value becomes dynamic. E.g., with the requirement reduced to 100: 100 XP on char 1 unlocks slot 2, then 100 on char 2 unlocks slot 3, etc.

### 4. [GMUDRealmSelectionMenu.cs:36](GreaterMUD.Module/Menu/GMUDRealmSelectionMenu.cs) — display effective value

In the realm loop, `long effectiveDupeExp = gmudServer.GetEffectiveDupeExpRequired();` and use it in the line-36 format (both the `> 0` check and the `"N0"` display). One in-memory scan per realm per render — negligible.

### 5. [greatermud.example.ini](GreaterMUD.Module/greatermud.example.ini) — document the settings

File uses bare `GMUDDev1_`-prefixed keys with no comments; `DupeExpRequired` isn't documented today. Add all three in that style:

```ini
GMUDDev1_DupeExpRequired=8000000000
GMUDDev1_DupeExpReduction=0
GMUDDev1_DupeExpMinimum=0
```

## Out of scope (noted)

- `GreaterMUD\` legacy project also has dupe-related code (SysCommand "showdupes") — that's IP-duplicate detection, unrelated; legacy project is reference-only.
- Settings are loaded once at startup (no reload mechanism); the new settings inherit that — only the qualifying **count** is live.

## Verification

1. `dotnet build TGS_5.sln` — clean compile (no automated tests exist in the solution).
2. **Regression (feature off):** no new INI keys → realm menu shows configured Dupe XP Requirement; slot gating unchanged.
3. **Manual telnet scenario** (127.0.0.1:2427), with small test values, e.g. `DupeExpRequired=1000`, `DupeExpReduction=200`, `DupeExpMinimum=300`:
   - 0 qualifying chars → realm menu shows 1,000; empty slot 2 shows "Exp needed: 1,000 − slot-1 XP".
   - Raise one character to ≥ 1000 XP → requirement displays 800; that account's slot 2 becomes creatable.
   - 4+ qualifying chars → 1000 − 800 = 200 < 300, so value clamps at 300.
   - Reroll a qualifying char → its `Rerolled` record stops counting; the `FromReroll` replacement counts only once its own XP ≥ 1000.
   - `DupeExpRequired=0` with nonzero Reduction/Minimum → realm menu shows "None", slot creation ungated.
   - **Existing-character protection:** create a character in slot 2 while the requirement is low, then make the effective requirement rise (reroll/delete the qualifying characters so the count drops). The slot-2 character must still be selectable; only creating into a new empty slot is gated.
