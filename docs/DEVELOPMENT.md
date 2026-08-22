# Draw Rounded Rectangle — development notes

The API findings behind this macro. `README.md` covers what it does and how to
use it.

## Geometry and the degrees-of-freedom budget

Four lines and four tangent arcs are drawn into the **active sketch**, centred on
its origin. Every relation is added explicitly — eight merged endpoints, eight
tangents, four horizontal/vertical, and three equal-radius — which leaves exactly
**five degrees of freedom**: position in x and y, plus width, height and radius.

Five is the right number, and it is the whole design. Drag a straight side and the
shape stays a rounded rectangle rather than coming apart; add width, height and
radius dimensions later and the sketch becomes fully defined with no redundancy.
Adding relations until it "looks defined" produces a shape that fights the user.

**No dimensions are added**, deliberately. Three inputs and three driving
dimensions proved slower in practice than typing one number and adjusting after.

## API traps worth keeping

- **`ISketchManager.AddToDB` must be `True`** before creating geometry, and
  restored afterwards. Left at its default, `SketchManager` draws through the user
  interface and **silently returns `Nothing` when geometry falls outside the
  visible graphics area**. This is the classic cause of a sketch macro that "used
  to work" — it depends on the zoom level, not the code.
- **The API is metres throughout.** `Dimension.SystemValue` is metres;
  `Dimension.Value` is user units. The form converts millimetres by dividing by
  1000.
- Relations are added by selecting entities and calling `SketchAddConstraints`;
  selection order matters and a wrong order fails silently rather than erroring.

## The icon

A single 128×128 RGBA image. **Thin strokes wash out on the toolbar** —
SOLIDWORKS renders toolbar icons at roughly 20 px, and a smooth downscale is
unkind to thin light strokes even when the source is fully opaque.

The fix that worked was **drawing the lines heavier in the one large image**.
Strokes with enough weight survive the resample, so no separate small version is
needed. Try that before anything more elaborate.

## Known limitations

- Requires a sketch already open for editing; it will not create one.
- The corner radius is fixed at one tenth of the scale.

## Verification status

Confirmed working in SOLIDWORKS. The relation behaviour in particular was checked
by dragging the result, which is the only way to establish it — no amount of
reading the code proves a sketch stays coherent under a drag.

## There is no build step

A `.swp` is a binary VBA project. Editing the `.vba` in `src\` changes nothing
that runs until the source is pasted into the SOLIDWORKS VBA editor and saved.
Treat the `.vba` as the source of truth and re-paste after every change.

Do not patch a `.swp` directly: it stores compiled p-code ahead of the source
text and VBA runs the p-code, so a patched file shows new code in the editor
while still running the old.
