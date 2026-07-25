---
id: images
title: Images and Pictures
sidebar_label: Images
description: Embed images in a worksheet, anchor them absolutely or to cells, scale them, control how they behave when rows and columns move, and group them.
---

# Images and Pictures

XLibur can embed images in a worksheet — a company logo on a report header, a product photo
beside each row, a chart rendered elsewhere and dropped in as a picture.

The picture types live in a **separate namespace**, so you need an extra using directive:

```csharp
using XLibur.Excel;
using XLibur.Excel.Drawings;   // IXLPicture, XLPictureFormat, XLPicturePlacement
```

## Adding a picture

From a file path — the format is inferred from the extension:

```csharp
using var workbook = new XLWorkbook();
var ws = workbook.Worksheets.Add("Report");

ws.AddPicture("logo.png")
  .MoveTo(ws.Cell("A1"));
```

From a stream, which is what you want for images that come from a database, an HTTP response,
or an embedded resource. State the format explicitly:

```csharp
await using var stream = File.OpenRead("logo.png");

ws.AddPicture(stream, XLPictureFormat.Png, "CompanyLogo")
  .MoveTo(ws.Cell("A1"));
```

```csharp
// From an embedded resource
var assembly = typeof(Program).Assembly;
using var resource = assembly.GetManifestResourceStream("MyApp.Assets.logo.png")!;

ws.AddPicture(resource, XLPictureFormat.Png, "Logo");
```

Supported formats: `Png`, `Jpeg`, `Gif`, `Bmp`, `Tiff`, `Emf`, `Wmf`, `Webp`, `Svg`, `Pcx`,
`Icon`.

:::note
`AddPicture` reads the stream immediately into an internal buffer, so you may dispose the
source stream straight away. Reusing one stream for several pictures means resetting it —
`stream.Position = 0` — between calls.
:::

## Anchoring

Excel supports three anchor styles, and XLibur picks between them based on which `MoveTo`
overload you use.

**Absolute** — pixel offsets from the top-left of the sheet. The picture ignores rows and
columns entirely:

```csharp
ws.AddPicture("logo.png").MoveTo(220, 150);   // left, top, in pixels
```

**One-cell** — anchored to a single cell, keeping its own size:

```csharp
ws.AddPicture("logo.png").MoveTo(ws.Cell("B3"));

// With a pixel offset inside that cell
ws.AddPicture("logo.png").MoveTo(ws.Cell("B3"), 20, 5);
```

**Two-cell** — stretched between two cells, so it resizes as the columns and rows do:

```csharp
ws.AddPicture("banner.png").MoveTo(ws.Cell("B2"), ws.Cell("H10"));

// With offsets at both corners
ws.AddPicture("banner.png")
  .MoveTo(ws.Cell("B2"), 20, 5, ws.Cell("H10"), 30, 10);
```

Reading the anchor back:

```csharp
var picture = ws.Pictures.Picture("CompanyLogo");

Console.WriteLine(picture.TopLeftCell.Address);
Console.WriteLine(picture.BottomRightCell.Address);
Console.WriteLine($"{picture.Left}, {picture.Top}");
```

## Placement — behaviour when cells move

`XLPicturePlacement` controls what happens when the user inserts a row or resizes a column
under the picture:

```csharp
picture.WithPlacement(XLPicturePlacement.MoveAndSize);   // move and resize with cells
picture.WithPlacement(XLPicturePlacement.Move);          // move, but keep its size
picture.WithPlacement(XLPicturePlacement.FreeFloating);  // ignore cells entirely

picture.Placement = XLPicturePlacement.Move;
```

| Placement | Excel's label |
|---|---|
| `MoveAndSize` | Move and size with cells |
| `Move` | Move but don't size with cells |
| `FreeFloating` | Don't move or size with cells |

A logo in a header usually wants `Move` (so it stays put relative to the sheet content, without
being distorted by a column resize); a background or watermark wants `FreeFloating`.

## Sizing and scaling

Sizes are in pixels:

```csharp
picture.Width = 240;
picture.Height = 80;
picture.WithSize(240, 80);

Console.WriteLine($"{picture.OriginalWidth} x {picture.OriginalHeight}");
```

Scaling is relative to the *current* size by default, or to the original size with the flag:

```csharp
picture.Scale(0.5);                            // half the current size
picture.Scale(0.2, relativeToOriginal: true);  // 20% of the original

picture.ScaleWidth(1.5);
picture.ScaleHeight(0.75);
```

Preserving the aspect ratio while fitting a maximum width:

```csharp
const int maxWidth = 300;

if (picture.Width > maxWidth)
{
    picture.Scale((double)maxWidth / picture.Width);
}
```

## Finding, copying, and deleting

```csharp
Console.WriteLine(ws.Pictures.Count);

foreach (var p in ws.Pictures)
{
    Console.WriteLine($"{p.Name}: {p.Format} {p.Width}x{p.Height} at {p.TopLeftCell.Address}");
}

if (ws.Pictures.TryGetPicture("CompanyLogo", out var logo))
{
    logo!.Scale(0.5);
}

// Duplicate on the same sheet, or copy to another
var copy = picture.Duplicate();
var onOtherSheet = picture.CopyTo(workbook.Worksheet("Summary"));

picture.Delete();
ws.Pictures.Delete("CompanyLogo");
```

Names must be unique within the worksheet; omit the name and XLibur generates one.

## Reading an image back out

`ImageStream` returns the stored bytes, which is how you extract images from a workbook someone
sent you:

```csharp
using var workbook = new XLWorkbook("Catalogue.xlsx");

foreach (var ws in workbook.Worksheets)
{
    foreach (var picture in ws.Pictures)
    {
        var extension = picture.Format.ToString().ToLowerInvariant();
        var path = Path.Combine("extracted", $"{picture.Name}.{extension}");

        Directory.CreateDirectory("extracted");

        using var file = File.Create(path);
        picture.ImageStream.Position = 0;
        picture.ImageStream.CopyTo(file);
    }
}
```

## Grouping

Two or more free-floating pictures can be combined into a group shape, so Excel treats them as
one object:

```csharp
var a = ws.AddPicture("badge.png").MoveTo(100, 100).WithPlacement(XLPicturePlacement.FreeFloating);
var b = ws.AddPicture("label.png").MoveTo(180, 100).WithPlacement(XLPicturePlacement.FreeFloating);

var group = ws.Pictures.Group(a, b);
```

Once a group exists you can add to and remove from it:

```csharp
await using var extra = File.OpenRead("star.png");
group.Add(extra, "Star");

group.Remove(b);

foreach (var member in group.Pictures)
{
    Console.WriteLine(member.Name);
}
```

Every picture knows whether it belongs to one:

```csharp
if (picture.IsInGroup)
{
    Console.WriteLine(picture.Group!.Pictures.Count());
}
```

:::note
`Pictures.Group` requires the pictures to already use free-floating placement — call
`MoveTo(left, top)` or `WithPlacement(XLPicturePlacement.FreeFloating)` first.
:::

## Images inside cells

Standard Excel pictures float *over* the grid rather than living in a cell, which is why they
do not sort or filter with the data. If you need a picture that behaves like a cell value,
Excel 365's "Place in Cell" images are a different feature — a floating picture positioned over
a cell is the closest equivalent in a `.xlsx` that older Excel versions will read.

## A worked example

```csharp
using XLibur.Excel;
using XLibur.Excel.Drawings;

using var workbook = new XLWorkbook();
var ws = workbook.Worksheets.Add("Products");

// Leave room for a logo band across the top
ws.Row(1).Height = 60;
ws.Range("A1:E1").Merge();

await using (var logo = File.OpenRead("assets/logo.png"))
{
    ws.AddPicture(logo, XLPictureFormat.Png, "Logo")
      .MoveTo(ws.Cell("A1"), 8, 8)
      .WithPlacement(XLPicturePlacement.Move)
      .Scale(0.4, relativeToOriginal: true);
}

// Header row
ws.Cell("A3").Value = "SKU";
ws.Cell("B3").Value = "Product";
ws.Cell("C3").Value = "Photo";
ws.Cell("D3").Value = "Price";
ws.Range("A3:D3").Style.Font.Bold = true;

ws.Column("C").Width = 16;

var products = new[]
{
    ("SKU-001", "Widget", "assets/widget.png", 9.99m),
    ("SKU-002", "Gadget", "assets/gadget.png", 24.50m),
};

var row = 4;
foreach (var (sku, name, imagePath, price) in products)
{
    ws.Cell(row, 1).Value = sku;
    ws.Cell(row, 2).Value = name;
    ws.Cell(row, 4).Value = price;
    ws.Row(row).Height = 70;

    await using var image = File.OpenRead(imagePath);
    var picture = ws.AddPicture(image, XLPictureFormat.Png, $"Photo_{sku}")
        .MoveTo(ws.Cell(row, 3), 4, 4)
        .WithPlacement(XLPicturePlacement.Move);

    // Fit within the row height, preserving aspect ratio
    const int maxHeight = 80;
    if (picture.Height > maxHeight)
    {
        picture.Scale((double)maxHeight / picture.Height);
    }

    row++;
}

ws.Range($"D4:D{row - 1}").Style.NumberFormat.Format = "$ #,##0.00";
workbook.SaveAs("ProductCatalogue.xlsx");
```

## Where to next

- [Charts](./charts.md) — the other drawing type, with the same anchor model
- [Worksheets](./worksheets.md) — row heights and column widths that pictures anchor against
