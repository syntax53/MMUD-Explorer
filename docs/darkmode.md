# Dark Mode — Developer Handoff

Everything about MME's optional dark mode: how it works, the rules new code
must follow, what was deliberately left alone, and what not to re-attempt.
The condensed rules also live in `CLAUDE.md`; this is the full reference.

## Quick facts

| | |
|---|---|
| Owner module | `modTheme.bas` (all theming logic lives here) |
| Setting | `[Settings] DarkMode=1` in `settings.ini` |
| Toggle UI | `chkDarkMode` in frmSettings |
| Mode switch | **Requires app restart** (by design) |
| Startup read | `Sub Main` (modMain) — resolved before any form loads |
| Light theme | The design-time `.frm` colors ARE the light theme; dark is applied over them at load |
| Menu bar | Owner-drawn dark via `WM_UAH*` in `modMenuSubClass.bas` (frmMain only) |
| Title bars | `DwmSetWindowAttribute` attr 20 (19 fallback), Win10+ |

## How it applies

Every form calls `ApplyDarkTheme Me` at the top of `Form_Load`. In light mode
it exits immediately; in dark mode it walks `frm.Controls` once and recolors
by control type. Forms are recreated from design-time state on every Load, so
re-loading a form (including the in-process `AppReload`) is safe.

**A new form must add the `ApplyDarkTheme Me` call — nothing else is needed
for standard controls.**

## The public API (modTheme.bas)

| Member | Use |
|---|---|
| `bDarkMode` | Global flag; read-only outside startup code |
| `ApplyDarkTheme frm` | Call once from every `Form_Load` |
| `TColor(color)` | Translate any color for the active theme. Light mode: pass-through. Dark: system colors map to the palette; explicit RGB gets lightness-inverted (hue kept, clamped off pure white/black) |
| `TBtnColor(color)` | Same idea for CommandButton faces only: remaps the default face to `DK_BTN_FACE`, passes custom highlight colors through |
| `SetFrameCaption fra, s` | Change a frame caption at runtime (updates the dark-mode overlay label and border gap) |
| `SetFrameForeColor fra, c` | Change a frame caption color at runtime (same reason) |
| `SetFrameFontBold fra, b` | Change a frame caption's bold state at runtime (re-measures the caption gap in the border line) |
| `ApplyDarkTitleBar hWnd` | Dark title bar for a window (already called by ApplyDarkTheme) |
| `DK_FORM_BACK` &H202020 | window/dialog background |
| `DK_FIELD_BACK` &H262525 | text/list/combo field background |
| `DK_TEXT` &HE0E0E0 | standard text |
| `DK_TEXT_DIM` &H909090 | disabled/gray text |
| `DK_BTN_FACE` &H989898 | button face (captions always black — VB6 limit) |
| `DK_LINE` &H505050 | separator lines / muted borders |

## Rules for new/changed UI code

1. **Runtime color assignments** must be wrapped: `x.ForeColor = TColor(&HC0&)`,
   button faces via `TBtnColor(...)`. Never wrap a value twice — inversion is
   an involution; the second pass restores the original. (`ColorListviewRow`
   wraps internally; its callers pass raw colors.)
2. **Color comparisons** against themed values must compare the same wrapped
   value — see the `TColor(&H80000008)` comparisons in `modListViewExt.bas`.
3. **Frame captions**: in dark mode the real caption is not drawn (BorderStyle
   is 0); an overlay label `lblDkFrameCap*` stands in. Any runtime
   `.Caption`/caption-`.ForeColor` change must go through `SetFrameCaption` /
   `SetFrameForeColor` or it silently won't display. Bold changes must go
   through `SetFrameFontBold` — the overlay shares the frame's Font object so
   the text itself bolds either way, but the caption gap in the border line
   must be re-measured or the line runs through the widened text.
4. **Startup state**: VB6 fires `chk_Click` only when a value *changes*. If a
   checkbox controls a themed color and its saved state equals the design-time
   state, initialize the color explicitly in `Form_Load` (see the
   `chkGlobalFilter` block in frmMain).
5. **New CommandButtons**: give them `Style = 1 'Graphical` in the designer so
   `BackColor` works. In this unthemed (no comctl6 manifest) app that renders
   identically in light mode.
6. **Opting a control out**: put `notheme` anywhere in its `Tag`, or give it a
   custom opaque `BackColor` (see next section).

## What gets themed, by control type

| Control | Dark-mode treatment |
|---|---|
| Form | back = DK_FORM_BACK; dark title bar |
| Label | fore/back themed **only if** back is a system color (or BackStyle transparent: fore only). `lblRoomCell` always skipped (map data; code compares its BackColor) |
| TextBox / ListBox | back → DK_FIELD_BACK (white/window) or inverted; sunken client edge swapped for thin border; **custom opaque back = skipped entirely** |
| ComboBox | same as TextBox + post-paint border overdraw via subclass (combos paint their sunken border internally) |
| CheckBox / OptionButton | themed only if back is a system color |
| Frame | system border dropped; muted group-box border drawn from subclass (top line at caption mid-height with caption gap, geometry packed in `dwRefData`); caption recreated as overlay label; **custom opaque back = skipped entirely** (keeps original border/caption) |
| CommandButton | default face → DK_BTN_FACE; custom faces kept (highlight colors); captions stay black (no ForeColor in VB6) |
| ListView / TreeView | back/fore themed; **gridlines turned off** (their color is baked into the control); sunken edge muted |
| Line | BorderColor → DK_LINE |
| PictureBox | back themed only if buttonface; sunken edge stripped |
| cntSplitter | back → DK_FORM_BACK (BackColor property added to the UserControl for this) |
| Shape / Image / Timer / Menu | untouched |

Pre-styled dark islands that rely on the custom-back skip: the black
char-stat panel (`lblInvenCharStat`/`lblInvenStats`), map room cells, map
legend swatches, frmMap's options panel, `txtMapMove` (frmMain and
frmMegaMUDPathing), frmMap's `txtRoomRoom`/`txtRoomMap`, frmAbout's text box.

## Subclassing details

- Uses comctl32 `SetWindowSubclass` (ordinals #410/#412/#413, same as
  modForms), gated by `gbAllowSubclassing`; unhooks on WM_DESTROY/WM_NCDESTROY.
- `DarkComboBorderProc` (id &H444B31): after WM_PAINT, overdraws the combo's
  2px sunken border ring (outer DK_LINE, inner DK_FIELD_BACK).
- `DarkFrameBorderProc` (id &H444B32): after WM_PAINT, draws the group-box
  border. `dwRefData` packs caption geometry: bits 0-7 = caption half-height
  (border line y), 8-15 = gap left px, 16-31 = gap right px; 0 = no caption
  (plain rectangle). `SetFrameCaption` re-packs when a caption changes.
- The frmMain menu bar: `WM_UAHDRAWMENU`/`WM_UAHDRAWMENUITEM` handling plus a
  bottom-line repaint on WM_NCPAINT/WM_NCACTIVATE, in `modMenuSubClass.bas`.
  Only active in builds where the wndproc subclass is installed (non-dev
  branch of frmMain's Form_Load).

## Intentionally NOT themed (and why)

- **Scrollbars, ListView column headers, popup/context menus, MsgBoxes** stay
  system-light. The uxtheme dark ordinals (`SetPreferredAppMode` #135,
  `FlushMenuThemes` #136, `RefreshImmersiveColorPolicyState` #104,
  `AllowDarkModeForWindow` #133, `SetWindowTheme "DarkMode_Explorer"`) were
  tried and **fully reverted**: they act process-wide (VB6.EXE itself when
  running from the IDE), destabilized the IDE (run-time error 97 at close),
  and visibly changed nothing in this unmanifested classic-rendered app.
  **Do not re-attempt without moving to a comctl6-manifested build, which
  would break the classic rendering the rest of the theme depends on.**
- ListView headers could in principle be dark via NM_CUSTOMDRAW owner-draw,
  but mscomctl's notification routing is unverified — parked by user decision.
- ListView gridlines: not recolorable on classic comctl; disabled in dark.

## Button icon spec

Buttons are `#F0F0F0` (light) / `#989898` (dark, `DK_BTN_FACE`). VB6 pictures
have 1-bit transparency (no alpha, no PNG).

- **Format**: 24-bit BMP, exact pixel size (16x16 / 24x24), tagged **96 DPI**
  (72-DPI files get scaled and blur), background pure magenta RGB(255,0,255).
- **Per button**: `Picture` = bmp, `UseMaskColor` = True,
  `MaskColor` = `&H00FF00FF&`.
- **Edges**: hard/aliased only; if antialiasing is unavoidable, pre-blend it
  toward `#C4C4C4` (midpoint of the two faces). Avoid pure white or very pale
  glyph colors — they vanish on the dark face.
- A generated clean set exists in `icons-new\` (same base names as the old
  GIFs); pictures embed into `.frx`, so assignment happens in the IDE only.

## Troubleshooting

| Symptom | Likely cause |
|---|---|
| A color looks light-mode-ish in dark (or vice versa) | Raw runtime assignment not wrapped in `TColor`/`TBtnColor` |
| A color is washed out / wrong hue in dark | Value passed through `TColor` twice |
| Frame caption text doesn't change / wrong color | Raw `.Caption`/`.ForeColor` set instead of `SetFrameCaption`/`SetFrameForeColor` |
| A control shows light text on light custom back | Its back is a *system* light color (add it to `TColor`'s map, like COLOR_MENU &H80000004 was) |
| Wrong color only at startup, fixes itself on toggle | Click-event init gap — rule 4 above |
| Combo/frame borders bright in a build | `gbAllowSubclassing` is False on that code path |
| Run-time error 97 closing the VB6 IDE | Something reintroduced process-wide uxtheme calls or an unremoved subclass |
