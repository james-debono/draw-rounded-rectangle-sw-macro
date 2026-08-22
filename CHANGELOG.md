# Changelog

Semantic Versioning. `MAJOR` reaches 1 when the behaviour is settled enough
to promise not to break it; 0.x is an honest statement that it may still move.

---

## 0.4.0 — 2026-08-21

- **Renamed from "Draw Squound" to "Draw Rounded Rectangle."** The old name
  was coined and told nobody what the macro did; the new one is what the shape
  is actually called. The form's title bar, the failure message and the public
  `DrawRoundedRectangle` procedure all follow.
- Moved to its own repository, with the `Source` URL updated to match.
- No functional change.

## 0.3.3 — 2026-08-20

- The version in the form's title bar is now shown in brackets — `Draw Squound
  (0.3.3)` — so it reads as metadata rather than part of the name.
- The `Source` URL in the header now points at the renamed repository,
  `draw-squound-sw-macro`.
- No functional change.

## 0.3.2 — 2026-08-13

- The version now appears in the form's title bar, so it's visible whenever the
  macro is used rather than only by opening the code. Set from a `MACRO_VERSION`
  constant that `build-library.ps1` checks against the header.
- Documentation wording changed from "size" back to **"scale"**, matching what the
  form actually says.
- No functional change.

## 0.3.1 — 2026-08-09

- Released under the **MIT licence**. The full licence text is now carried in the
  code itself, so a `.swp` passed on by itself still carries its licence.
- Header rewritten: what the macro does and how the relations behave, plus
  version, date, author and source. Maintenance notes moved below
  `Option Explicit`.
- No functional change.

## 0.3.0 — 2026-08-01

- One "Scale" input instead of three, and no driving dimensions. Typing one
  number and adjusting the shape afterwards proved faster than entering three
  values and then having to delete the dimensions to change anything.
- `DrawAndDimensionSketch` renamed to `DrawSquound`, since it no longer
  dimensions anything.

Confirmed working in SOLIDWORKS on 2026-08-09.

## 0.2.1 — 2026-07-30

- Draws into the sketch currently open for editing, rather than creating a new
  sketch on the Front Plane. Selecting a reference plane was failing, and
  drawing onto whatever sketch is already open is what was wanted.

## 0.2.0 — 2026-07-30

- Complete rewrite of the drawing routine.

  0.1.0 was written against an API that does not exist — `SketchManager` has no
  `AddConstraint`, there is no `swSketchCONSTRAINTTYPE_e` enum, and
  `DisplayDimension` has neither `.Value` nor `.Name`. Every one of those is a
  compile error, which is why it never ran. The rewrite verifies each call
  against the installed type libraries.

## 0.1.0 — 2025-10-29

- Original version. Never ran; see 0.2.0.
