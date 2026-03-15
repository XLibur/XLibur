using System;

namespace XLibur.Excel;

internal sealed class XLSparklineVerticalAxis : IXLSparklineVerticalAxis
{
    #region Public Properties

    public double? ManualMax
    {
        get => _manualMax;
        set => SetManualMax(value);
    }

    public double? ManualMin
    {
        get => _manualMin;
        set => SetManualMin(value);
    }

    public XLSparklineAxisMinMax MaxAxisType
    {
        get => _maxAxisType;
        set => SetMaxAxisType(value);
    }

    public XLSparklineAxisMinMax MinAxisType
    {
        get => _minAxisType;
        set => SetMinAxisType(value);
    }

    #endregion Public Properties

    #region Public Methods

    public IXLSparklineVerticalAxis SetManualMax(double? value)
    {
        if (value != null)
            MaxAxisType = XLSparklineAxisMinMax.Custom;

        _manualMax = value;
        return this;
    }

    public IXLSparklineVerticalAxis SetManualMin(double? value)
    {
        if (value != null)
            MinAxisType = XLSparklineAxisMinMax.Custom;

        _manualMin = value;
        return this;
    }

    public IXLSparklineVerticalAxis SetMaxAxisType(XLSparklineAxisMinMax value)
    {
        if (value != XLSparklineAxisMinMax.Custom)
            _manualMax = null;

        _maxAxisType = value;
        return this;
    }

    public IXLSparklineVerticalAxis SetMinAxisType(XLSparklineAxisMinMax value)
    {
        if (value != XLSparklineAxisMinMax.Custom)
            _manualMin = null;

        _minAxisType = value;
        return this;
    }

    #endregion Public Methods

    #region Private Fields

    private double? _manualMax;
    private double? _manualMin;
    private XLSparklineAxisMinMax _maxAxisType;
    private XLSparklineAxisMinMax _minAxisType;

    #endregion Private Fields

    public IXLSparklineGroup SparklineGroup { get; }

    public XLSparklineVerticalAxis(IXLSparklineGroup sparklineGroup)
    {
        SparklineGroup = sparklineGroup ?? throw new ArgumentNullException(nameof(sparklineGroup));
    }

    public static void Copy(IXLSparklineVerticalAxis from, IXLSparklineVerticalAxis to)
    {
        to.ManualMax = from.ManualMax;
        to.ManualMin = from.ManualMin;
        to.MaxAxisType = from.MaxAxisType;
        to.MinAxisType = from.MinAxisType;
    }
}
