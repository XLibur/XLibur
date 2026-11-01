using System;
using DocumentFormat.OpenXml;
using A = DocumentFormat.OpenXml.Drawing;

namespace XLibur.Excel.Drawings;

/// <summary>The DrawingML preset geometry and default name of each <see cref="XLShapeType"/>.</summary>
internal static class XLShapeTypes
{
    private static readonly (XLShapeType Type, string Preset, string Name)[] Presets =
    [
        (XLShapeType.TextBox, "rect", "TextBox"),
        (XLShapeType.Rectangle, "rect", "Rectangle"),
        (XLShapeType.RoundedRectangle, "roundRect", "Rounded Rectangle"),
        (XLShapeType.Ellipse, "ellipse", "Oval"),
        (XLShapeType.Triangle, "triangle", "Isosceles Triangle"),
        (XLShapeType.RightArrow, "rightArrow", "Right Arrow"),
        (XLShapeType.LeftArrow, "leftArrow", "Left Arrow"),
        (XLShapeType.UpArrow, "upArrow", "Up Arrow"),
        (XLShapeType.DownArrow, "downArrow", "Down Arrow"),
        (XLShapeType.Line, "line", "Line"),
        (XLShapeType.RoundedCallout, "wedgeRoundRectCallout", "Speech Bubble: Rectangle with Corners Rounded"),
        (XLShapeType.WedgeRectCallout, "wedgeRectCallout", "Speech Bubble: Rectangle"),
        (XLShapeType.Diamond, "diamond", "Diamond"),
        (XLShapeType.Pentagon, "pentagon", "Pentagon"),
        (XLShapeType.Chevron, "chevron", "Chevron"),
        (XLShapeType.Star5, "star5", "Star: 5 Points"),
        (XLShapeType.Can, "can", "Cylinder"),
        (XLShapeType.Cloud, "cloud", "Cloud"),
    ];

    /// <summary>The preset geometry of <paramref name="type"/>.</summary>
    /// <exception cref="ArgumentException"><paramref name="type"/> is <see cref="XLShapeType.Custom"/> or not a member.</exception>
    internal static string PresetOf(XLShapeType type)
    {
        foreach (var (t, preset, _) in Presets)
        {
            if (t == type)
                return preset;
        }

        throw new ArgumentException(type == XLShapeType.Custom
            ? "A custom shape has no preset of its own. Use AddPreset with the geometry's name."
            : $"Unknown shape type '{type}'.", nameof(type));
    }

    /// <summary>The type a geometry reads back as: a text box is a rectangle with <c>txBox</c> set.</summary>
    internal static XLShapeType FromPreset(string preset, bool isTextBox)
    {
        if (isTextBox && preset == "rect")
            return XLShapeType.TextBox;

        foreach (var (type, p, _) in Presets)
        {
            if (type != XLShapeType.TextBox && p == preset)
                return type;
        }

        return XLShapeType.Custom;
    }

    /// <summary>The start of the name Excel gives a shape of the geometry, such as <c>"Oval"</c>.</summary>
    internal static string NameOf(string preset, bool isTextBox)
    {
        var type = FromPreset(preset, isTextBox);
        foreach (var (t, _, name) in Presets)
        {
            if (t == type)
                return name;
        }

        return "Shape";
    }

    /// <summary>Whether <paramref name="preset"/> is a DrawingML preset geometry, <c>ST_ShapeType</c>.</summary>
    internal static bool IsPreset(string preset) => ((IEnumValue)new A.ShapeTypeValues(preset)).IsValid;
}
