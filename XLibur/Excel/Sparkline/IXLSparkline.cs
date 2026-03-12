namespace XLibur.Excel;

public interface IXLSparkline
{
    #region Public Properties

    bool IsValid { get; }

    IXLCell Location { get; set; }

    IXLRange SourceData { get; set; }

    IXLSparklineGroup SparklineGroup { get; }

    #endregion Public Properties

    #region Public Methods

    IXLSparkline SetLocation(IXLCell value);

    IXLSparkline SetSourceData(IXLRange value);

    #endregion Public Methods
}
