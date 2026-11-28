namespace XLibur.Excel;

public enum XLDrawingAnchor { MoveAndSizeWithCells, MoveWithCells, Absolute }
public interface IXLDrawing<out T>
{
    int ShapeId { get; }

    bool Visible { get; set; }
    T SetVisible();
    T SetVisible(bool hidden);

    ////String Name { get; set; }
    ////T SetName(String name);

    ////String Description { get; set; }
    ////T SetDescription(String description);

    IXLDrawingPosition Position { get; }

    int ZOrder { get; set; }
    T SetZOrder(int zOrder);

    IXLDrawingStyle Style { get; }

}
