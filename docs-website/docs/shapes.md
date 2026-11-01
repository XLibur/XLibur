---
id: shapes
title: Shapes and Text Boxes
sidebar_label: Shapes
description: Add text boxes, callouts, arrows and other preset shapes to a worksheet, format their fill, line and text, and edit the shapes of a file Excel wrote without losing their effects.
---

# Shapes and Text Boxes

XLibur can add drawing shapes to a worksheet: a floating text box, a callout that points at a
figure, an arrow, a star. It also reads the shapes of a file Excel wrote, so you can change their
text and formatting.

The shape types are in the drawings namespace:

```csharp
using XLibur.Excel;
using XLibur.Excel.Drawings;   // IXLShape, IXLTextBox, XLShapeType
```

## Adding a text box

A text box is Excel's own: a rectangle with a white fill and a thin line, as Insert ▸ Text Box makes.

```csharp
using var workbook = new XLWorkbook();
var ws = workbook.Worksheets.Add("Report");

ws.Shapes.AddTextBox("Draft — do not circulate", ws.Cell("B2"))
         .WithSize(220, 60);
```

Sizes and positions are in pixels, as for pictures. A text box added at a cell moves with that cell
and keeps its size. Give pixel coordinates instead to place it on the sheet itself:

```csharp
ws.Shapes.AddTextBox("Floating note", left: 300, top: 40);
```

## Adding other shapes

`Add` takes a preset type. The new shape looks as it does when you add it from Excel's ribbon: it
takes its fill, line and font from the workbook's theme.

```csharp
var callout = ws.Shapes.Add(XLShapeType.WedgeRectCallout, ws.Cell("E4"));
callout.FillColor = XLColor.LightYellow;
callout.LineColor = XLColor.Black;
callout.LineWidthPt = 1.5;
callout.LineDash = XLLineDashStyle.Dash;
```

`XLShapeType` names the common shapes: rectangles, rounded rectangles, ellipses, triangles,
arrows, a line, two callouts, a diamond, a pentagon, a chevron, a star, a can and a cloud.
`AddPreset` adds any other of DrawingML's preset geometries by name:

```csharp
var moon = ws.Shapes.AddPreset("moon", ws.Cell("H2"));
// moon.ShapeType is XLShapeType.Custom; moon.PresetGeometry is "moon".
```

## Placement

A shape uses the same `XLPicturePlacement` as a picture:

| Placement | Added with | Behaviour |
|---|---|---|
| `Move` | `MoveTo(cell)`, `MoveTo(cell, x, y)` | Moves with its top-left cell and keeps its size. The default. |
| `MoveAndSize` | `MoveTo(fromCell, toCell)` | Moves and sizes with its cells. `Width` and `Height` cannot be set. |
| `FreeFloating` | `MoveTo(left, top)` | Stays where it is on the sheet. |

A shape loaded from a file can be given a new placement: it moves into an anchor of the new kind
with all its formatting.

## Text

A shape's text is a list of paragraphs, and each paragraph is a list of runs. A run is text in one
font.

```csharp
var body = callout.TextBody;
body.VerticalAlignment = XLTextVerticalAlignment.Center;

var line = body.AddParagraph();
line.Alignment = XLTextAlignment.Center;
line.AddText("Revised ").SetBold().SetFontColor(XLColor.DarkRed);
line.AddText("2026-08-01");

body.AddParagraph("Second line");
```

`IXLTextBox.Text` and `IXLTextBody.Text` read the whole text, with the paragraphs joined by `\n`.
Setting either replaces every paragraph. The new paragraphs take the alignment of the first
paragraph there was, and the new runs the font of its first run. A line break inside a paragraph
reads as ``; to write one, start a new paragraph. Control characters other than a tab are
refused when you set the text, because the file cannot hold them.

The font properties of a run are `Bold`, `Italic`, `Underline`, `Strikethrough`, `FontSize`,
`FontColor` and `FontName`. Each is nullable: `null` means not set, so the run uses the shape's
style.

## Editing shapes in a file

Shapes in a loaded workbook are in `ws.Shapes`, by name:

```csharp
using var workbook = new XLWorkbook("report.xlsx");
var ws = workbook.Worksheet(1);

var note = (IXLTextBox)ws.Shapes.Shape("TextBox 1");
note.TextBody[0][0].Text = "Approved";   // changes only that run's text
ws.Shapes.Shape("Arrow").LineWidthPt = 3; // keeps the arrowheads
workbook.Save();
```

XLibur writes back only what you change, into the shape's own XML. What it does not model
survives the save: gradients, shadows, glow, 3-D, rotation, hyperlinks, adjust handles. A shape you
do not change is not written at all.

A formatting property reads `null` when the file sets it in a way the model does not hold, such as
a gradient fill, or when the colour comes only from the shape's theme style. Setting it writes an
explicit value. `NoFill` and `NoLine` remove a fill or a line.

Names do not have to be unique, because Excel lets two shapes share one. `Shape(name)` returns the
first in drawing order, and `Delete(name)` deletes all of them. To tell such shapes apart, compare
their `Id`.

## Limits

- Shapes inside a group, connectors, and shapes with a custom geometry are kept as the file has
  them, but are not in `ws.Shapes`.
- Rotation, flips, auto-fit, text insets, vertical text, gradient fills, effects, bullets and
  hyperlinks are not in the model. A shape keeps them when you edit something else.
- Nothing measures the text: a shape is the size you give it, and long text overflows or wraps.
- Shapes do not move when rows or columns are inserted or deleted, and copying a worksheet does
  not copy them.
