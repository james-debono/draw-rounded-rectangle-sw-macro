# Draw Rounded Rectangle

A SOLIDWORKS macro that draws a rounded rectangle into the sketch you already have
open. Type one number for the scale, click Draw, and you get the shape with the
sketch relations that keep it a proper rounded rectangle.

Works with SOLIDWORKS 2022, 2024 and 2025.

## What you get

- The shape is drawn into **the sketch currently open for editing**, centred on
  that sketch's origin. Whichever plane or face the sketch sits on is where the
  shape lands.
- Corner radius is **one tenth** of the scale you type.
- **No dimensions are added.** The shape comes out under-defined on purpose, so
  you can drag it or dimension it yourself afterwards.

That last point is the reason it exists. Three inputs and three driving dimensions
turned out slower than typing one number and adjusting after.

## Relations, and why it stays a rounded rectangle

Every relation is added explicitly — eight merged endpoints, eight tangents, four
horizontal/vertical, and three equal-radius. That leaves exactly **five degrees of
freedom**: position in x and y, plus width, height and radius.

Five is the right answer. Drag a straight side and the shape stays a rounded
rectangle instead of coming apart. Add your own width, height and radius
dimensions later and the sketch becomes fully defined with no redundancy.

## Install

**The macro on its own:** download `Draw-Rounded-Rectangle.swp` from the
[latest release](../../releases/latest), then run it with **Tools > Macro > Run**,
or add it to a toolbar with **Tools > Customize > Commands > Macro**.

**With [MacroDeck](https://github.com/james-debono/macrodeck-sw-addin):** get the
[MacroDeck Collection](https://github.com/james-debono/macrodeck-collection-sw-macro-library/releases/latest),
which packages this macro with its icon and hover text alongside every other macro
in the set. Point MacroDeck at the unzipped folder and it appears as a button.

## Using it

1. Open or start a sketch.
2. Run the macro.
3. Type a scale in millimetres and click **Draw**.

If no document is open, or no sketch is being edited, it says so and stops rather
than asking for a scale it can't use.

## Building from source

`src\Draw-Rounded-Rectangle.vba` is the standard module and
`src\Draw-Rounded-Rectangle.Form.vba` is the form's code-behind. A `.swp` is a
binary VBA project, so it can only be produced from inside SOLIDWORKS — there is
no build step:

1. Open the `.swp` via **Tools > Macro > Edit**.
2. Paste the module source into the module, and the form source into the form's
   code window (right-click the form > View Code).
3. Save.

The **form layout** exists only inside the `.swp` and has no text source. Don't
import the form source as a `.frm` — that would overwrite the layout. And don't try
to patch the `.swp` directly: it stores compiled p-code ahead of the source, so a
patched file shows new code in the editor while still running the old.

Technical detail is in [docs/DEVELOPMENT.md](docs/DEVELOPMENT.md).

## Licence

MIT — see [LICENSE](LICENSE). Free to use, modify and share. The full licence text
is also carried inside the macro itself, so a `.swp` passed on by itself still
carries its licence.

Written by James Debono, with AI assistance. Everything here was tested by
hand in SOLIDWORKS — nothing that touches the API can be verified any other way.

## Trademarks

SOLIDWORKS is a registered trademark of Dassault Systèmes SolidWorks Corporation.
This project is independent: it is not affiliated with, endorsed by, or sponsored
by Dassault Systèmes, and uses only the published SOLIDWORKS API.
