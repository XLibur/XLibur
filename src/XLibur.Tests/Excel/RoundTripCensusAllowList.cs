using System;
using System.Collections.Generic;
using System.Linq;

namespace XLibur.Tests.Excel;

internal static class RoundTripCensusAllowList
{
    private const string ContentTypes = "[Content_Types]";
    private const string CustomProperties = "custom-properties";
    private const string Drawing = "drawing";
    private const string ExtendedProperties = "extended-properties";
    private const string MsExcelRdrichvalue = "ms-excel.rdrichvalue";
    private const string MsExcelRdrichvaluestructure = "ms-excel.rdrichvaluestructure";
    private const string MsExcelRdrichvaluetypes = "ms-excel.rdrichvaluetypes";
    private const string MsExcelRichValueRel = "ms-excel.richValueRel";
    private const string PackageRelationships = "package.relationships";
    private const string SpreadsheetmlCalcChain = "spreadsheetml.calcChain";
    private const string SpreadsheetmlComments = "spreadsheetml.comments";
    private const string SpreadsheetmlPivotCacheDefinition = "spreadsheetml.pivotCacheDefinition";
    private const string SpreadsheetmlPivotTable = "spreadsheetml.pivotTable";
    private const string SpreadsheetmlSharedStrings = "spreadsheetml.sharedStrings";
    private const string SpreadsheetmlSheetMain = "spreadsheetml.sheet.main";
    private const string SpreadsheetmlSheetMetadata = "spreadsheetml.sheetMetadata";
    private const string SpreadsheetmlStyles = "spreadsheetml.styles";
    private const string SpreadsheetmlTable = "spreadsheetml.table";
    private const string SpreadsheetmlWorksheet = "spreadsheetml.worksheet";
    private const string VmlDrawing = "vmlDrawing";

    private const string SchemaDefault = "the value is the default, so the attribute says nothing and is not written";
    private const string FalseToggle = "a toggle set to false, or to its \"none\" value, says the same as no toggle";
    private const string TrueToggle = "a toggle set to true is written as the bare element, which means true";
    private const string OutOfRange = "0 is outside the attribute's range, and Excel reads the attribute as absent";
    private const string RevisionUid = "a revision id (xr:uid) only serves Excel's change tracking, and the save does not keep it";
    private const string Ignorable = "mc:Ignorable names the namespaces of markup the save no longer writes, such as revision ids";
    private const string Count = "a count attribute restates how many children follow, and the census counts the children";
    private const string StylesPruned = "a style record nothing refers to is removed, a duplicate is merged, and every reference is renumbered";
    private const string StyleIndex = "a style index is renumbered when the styles are pruned, and dropped when it is the default; the c@s->numFmt, c@s->font and other c@s-> keys compare the style each cell resolves to";
    private const string InlineString = "an inline string is written to the shared string table, or not at all when it is empty; the census counts its text either way";
    private const string FormulaText = "t=\"str\" marks the text result of a formula; the text of a cell with no formula is written as a plain string";
    private const string Deprecated = "shapeId is deprecated, and Excel ignores it";
    private const string NothingToSay = "a cell or row with nothing else to say is not written; its value, formula, style and row attributes have keys of their own";
    private const string EmptyPart = "an empty part is removed, with its relationship, its content type and the reference to it";
    private const string ColumnsMerged = "adjacent columns with the same width and style are written as one range";
    private const string FormatsMerged = "conditional formats with the same rule are written as one over the union of their ranges (SaveOptions.ConsolidateConditionalFormatRanges)";
    private const string ValidationsMerged = "data validations with the same rule are written as one (SaveOptions.ConsolidateDataValidationRanges), and a rule moves between the main and the x14 form as its formula needs";
    private const string SharedFormulas = "a shared formula is written as a formula in each cell";
    private const string Volatile = "ca and aca mark a volatile formula, which Excel works out from the formula when it loads the file";
    private const string CalcChain = "the calculation chain is a cache, rebuilt from the formulas the save writes";
    private const string Spans = "spans is an optional hint, written only for the rows the save writes from the model";
    private const string DyDescent = "x14ac:dyDescent is a layout hint Excel recomputes from the row's fonts";
    private const string PreserveSpace = "xml:space=\"preserve\" is written only where the text starts or ends with whitespace";
    private const string EmptyText = "an empty element says nothing";
    private const string ReservedName = "a stale or duplicate _xlnm._FilterDatabase is regenerated from the autofilter, and a print area or print titles whose reference is invalid is dropped";
    private const string HiddenActiveSheet = "a hidden sheet that would be the active one is made visible on load (XLWorksheetTests.FirstSheetIsActive_WhenNotSpecified)";
    private const string UnusedString = "a shared string no cell uses is not written";
    private const string RebuiltContentTypes = "the content types list is rebuilt for the parts the package holds";
    private const string LegacyCellImage = "a cell image XLibur wrote before #739 is rewritten in the form Excel writes";
    private const string FirstPageNumber = "useFirstPageNumber is written with firstPageNumber; without one it means what no attribute means";
    private const string TwoDimensionalTable = "dtr applies only to a one-dimensional data table";
    private const string CommentShape = "comment shapes are written from the comments, so a comment shape with no comment is removed; other VML shapes are kept as they are";

    private const string LostACellImageIsWrittenAs = "a cell image is written as its rich value index, not as the #VALUE! Excel writes";
    private const string LostACellSNamedCellStyle = "a cell's named cell style is replaced by Normal, mostly when the sheet is written from the model";
    private const string LostASolidFillInTheSystem = "a solid fill in the system background colour (indexed 65) written as no fill when the sheet is written from the model";
    private const string LostCachedValueOfAFormulaXLibur = "cached value of a formula XLibur cannot evaluate (#NAME?) dropped when the sheet is written";
    private const string LostCommentPropertiesCommentPrAndItsAnchor = "comment properties (commentPr and its anchor) dropped";
    private const string LostCommentTextBoxAlignmentVMLDiv = "comment text box alignment (VML div style) dropped";
    private const string LostCommentTextRunCharsetAndIndexed = "comment text run charset and indexed colour dropped";
    private const string LostHyperlinkDisplayTextDropped = "hyperlink display text dropped";
    private const string LostPivotCacheItemFlagsUContainsNonDate = "pivot cache item flags (u, containsNonDate) dropped";
    private const string LostRebuiltPivotFormatDxfsDropReadingOrder = "rebuilt pivot format dxfs drop readingOrder and the vertical and horizontal border sides";
    private const string LostRichTextRunCharsetDropped = "rich text run charset dropped";
    private const string LostRichValueDetailOfAVALUE = "rich-value detail of a #VALUE! error (subType, argument) dropped";
    private const string LostRowStyleCustomFormatSDroppedWhen = "row style (customFormat, s) dropped when the sheet is written from the model";
    private const string LostRowThickBotAndThickTopDroppedWhen = "row thickBot and thickTop dropped when the sheet is written from the model";
    private const string LostSelectionOfAFrozenPaneDropped = "selection of a frozen pane dropped and activePane reset";
    private const string LostTableColumnDataFormatDataDxfIdDropped = "table column data format (dataDxfId) dropped";
    private const string LostTheFontFamilyAndInOne = "the font family, and in one file the font name and size, dropped from a cell's font when the sheet is written from the model";

    private static readonly AllowedLoss[] Noise =
    [
        new(SpreadsheetmlPivotTable, "pivotTableDefinition@useAutoFormatting=false", SchemaDefault),
        new(SpreadsheetmlPivotTable, "pivotTableStyleInfo@showColStripes=false", SchemaDefault),
        new(SpreadsheetmlPivotTable, "pivotTableStyleInfo@showLastColumn=false", SchemaDefault),
        new(SpreadsheetmlPivotTable, "pivotTableStyleInfo@showRowStripes=false", SchemaDefault),
        new(SpreadsheetmlSheetMain, "definedName@hidden=false", SchemaDefault),
        new(SpreadsheetmlSheetMain, "sheet@state=visible", SchemaDefault),
        new(SpreadsheetmlSheetMain, "workbookPr@date1904=false", SchemaDefault),
        new(SpreadsheetmlWorksheet, "c@s=0", SchemaDefault),
        new(SpreadsheetmlWorksheet, "c@t=n", SchemaDefault),
        new(SpreadsheetmlWorksheet, "col@outlineLevel=0", SchemaDefault),
        new(SpreadsheetmlWorksheet, "customFilter@operator=equal", SchemaDefault),
        new(SpreadsheetmlWorksheet, "f@aca=false", SchemaDefault),
        new(SpreadsheetmlWorksheet, "f@dt2D=false", SchemaDefault),
        new(SpreadsheetmlWorksheet, "f@dtr=false", SchemaDefault),
        new(SpreadsheetmlWorksheet, "pageSetup@fitToHeight=1", SchemaDefault),
        new(SpreadsheetmlWorksheet, "pageSetup@fitToWidth=1", SchemaDefault),
        new(SpreadsheetmlWorksheet, "pane@xSplit=0", SchemaDefault),
        new(SpreadsheetmlWorksheet, "row@collapsed=false", SchemaDefault),
        new(SpreadsheetmlWorksheet, "row@customFormat=false", SchemaDefault),
        new(SpreadsheetmlWorksheet, "row@customHeight=false", SchemaDefault),
        new(SpreadsheetmlWorksheet, "row@hidden=false", SchemaDefault),
        new(SpreadsheetmlWorksheet, "row@outlineLevel=0", SchemaDefault),
        new(SpreadsheetmlWorksheet, "selection@activeCellId=0", SchemaDefault),
        new(SpreadsheetmlWorksheet, "sheetFormatPr@outlineLevelCol=0", SchemaDefault),
        new(SpreadsheetmlWorksheet, "sheetFormatPr@outlineLevelRow=0", SchemaDefault),
        new(VmlDrawing, "v:fill@angle=0", SchemaDefault),
        new(SpreadsheetmlWorksheet, "b[@val=false]", FalseToggle),
        new(SpreadsheetmlWorksheet, "i[@val=false]", FalseToggle),
        new(SpreadsheetmlWorksheet, "strike[@val=false]", FalseToggle),
        new(SpreadsheetmlWorksheet, "u[@val=none]", FalseToggle),
        new(SpreadsheetmlSharedStrings, "b@val=true", TrueToggle),
        new(SpreadsheetmlWorksheet, "pageSetup@copies=0", OutOfRange),
        new(SpreadsheetmlWorksheet, "pageSetup@horizontalDpi=0", OutOfRange),
        new(SpreadsheetmlWorksheet, "pageSetup@scale=0", OutOfRange),
        new(SpreadsheetmlWorksheet, "pageSetup@verticalDpi=0", OutOfRange),
        new(SpreadsheetmlComments, "comment@xr:uid", RevisionUid),
        new(SpreadsheetmlPivotTable, "pivotTableDefinition@xr:uid", RevisionUid),
        new(SpreadsheetmlTable, "autoFilter@xr:uid", RevisionUid),
        new(SpreadsheetmlWorksheet, "autoFilter@xr:uid", RevisionUid),
        new(SpreadsheetmlWorksheet, "dataValidation@xr:uid", RevisionUid),
        new(SpreadsheetmlWorksheet, "hyperlink@xr:uid", RevisionUid),
        new(SpreadsheetmlWorksheet, "x14:dataValidation@xr:uid", RevisionUid, 34),
        new(MsExcelRdrichvaluetypes, "{http://schemas.microsoft.com/office/spreadsheetml/2017/richdata2}rvTypesInfo@mc:Ignorable", Ignorable),
        new(SpreadsheetmlComments, "comments@mc:Ignorable", Ignorable),
        new(SpreadsheetmlPivotTable, "pivotTableDefinition@mc:Ignorable", Ignorable),
        new(MsExcelRdrichvaluestructure, "{http://schemas.microsoft.com/office/spreadsheetml/2017/richdata2}rvStructures@count", Count),
        new(SpreadsheetmlPivotCacheDefinition, "page@count", Count),
        new(SpreadsheetmlPivotCacheDefinition, "pages@count", Count),
        new(SpreadsheetmlPivotCacheDefinition, "rangeSets@count", Count),
        new(SpreadsheetmlSharedStrings, "sst@count", Count),
        new(SpreadsheetmlSharedStrings, "sst@uniqueCount", Count),
        new(SpreadsheetmlStyles, "dxfs@count", Count),
        new(SpreadsheetmlWorksheet, "dataValidations@count", Count),
        new(SpreadsheetmlWorksheet, "x14:dataValidations@count", Count, 4),
        new(SpreadsheetmlStyles, "alignment", StylesPruned),
        new(SpreadsheetmlStyles, "alignment@horizontal", StylesPruned),
        new(SpreadsheetmlStyles, "alignment@indent", StylesPruned),
        new(SpreadsheetmlStyles, "alignment@justifyLastLine", StylesPruned),
        new(SpreadsheetmlStyles, "alignment@shrinkToFit", StylesPruned),
        new(SpreadsheetmlStyles, "alignment@textRotation", StylesPruned),
        new(SpreadsheetmlStyles, "alignment@vertical", StylesPruned),
        new(SpreadsheetmlStyles, "alignment@wrapText", StylesPruned),
        new(SpreadsheetmlStyles, "b", StylesPruned),
        new(SpreadsheetmlStyles, "b@val", StylesPruned),
        new(SpreadsheetmlStyles, "bgColor", StylesPruned),
        new(SpreadsheetmlStyles, "bgColor@auto", StylesPruned),
        new(SpreadsheetmlStyles, "bgColor@indexed", StylesPruned),
        new(SpreadsheetmlStyles, "bgColor@rgb", StylesPruned),
        new(SpreadsheetmlStyles, "bgColor@theme", StylesPruned),
        new(SpreadsheetmlStyles, "bgColor@tint", StylesPruned),
        new(SpreadsheetmlStyles, "border", StylesPruned),
        new(SpreadsheetmlStyles, "border@diagonalDown", StylesPruned),
        new(SpreadsheetmlStyles, "border@diagonalUp", StylesPruned),
        new(SpreadsheetmlStyles, "bottom", StylesPruned),
        new(SpreadsheetmlStyles, "bottom@style", StylesPruned),
        new(SpreadsheetmlStyles, "charset", StylesPruned),
        new(SpreadsheetmlStyles, "charset@val", StylesPruned),
        new(SpreadsheetmlStyles, "color", StylesPruned),
        new(SpreadsheetmlStyles, "color@auto", StylesPruned),
        new(SpreadsheetmlStyles, "color@indexed", StylesPruned),
        new(SpreadsheetmlStyles, "color@theme", StylesPruned),
        new(SpreadsheetmlStyles, "color@tint", StylesPruned),
        new(SpreadsheetmlStyles, "condense", StylesPruned),
        new(SpreadsheetmlStyles, "condense@val", StylesPruned),
        new(SpreadsheetmlStyles, "diagonal", StylesPruned),
        new(SpreadsheetmlStyles, "diagonal@style", StylesPruned),
        new(SpreadsheetmlStyles, "dxf", StylesPruned),
        new(SpreadsheetmlStyles, "dxfs", StylesPruned),
        new(SpreadsheetmlStyles, "extend", StylesPruned),
        new(SpreadsheetmlStyles, "extend@val", StylesPruned),
        new(SpreadsheetmlStyles, "family", StylesPruned),
        new(SpreadsheetmlStyles, "family@val", StylesPruned),
        new(SpreadsheetmlStyles, "fgColor", StylesPruned),
        new(SpreadsheetmlStyles, "fgColor@indexed", StylesPruned),
        new(SpreadsheetmlStyles, "fgColor@rgb", StylesPruned),
        new(SpreadsheetmlStyles, "fgColor@theme", StylesPruned),
        new(SpreadsheetmlStyles, "fgColor@tint", StylesPruned),
        new(SpreadsheetmlStyles, "fill", StylesPruned),
        new(SpreadsheetmlStyles, "font", StylesPruned),
        new(SpreadsheetmlStyles, "i", StylesPruned),
        new(SpreadsheetmlStyles, "i@val", StylesPruned),
        new(SpreadsheetmlStyles, "left", StylesPruned),
        new(SpreadsheetmlStyles, "left@style", StylesPruned),
        new(SpreadsheetmlStyles, "name", StylesPruned),
        new(SpreadsheetmlStyles, "name@val", StylesPruned),
        new(SpreadsheetmlStyles, "numFmt", StylesPruned),
        new(SpreadsheetmlStyles, "numFmt@formatCode", StylesPruned),
        new(SpreadsheetmlStyles, "numFmt@numFmtId", StylesPruned),
        new(SpreadsheetmlStyles, "outline", StylesPruned),
        new(SpreadsheetmlStyles, "outline@val", StylesPruned),
        new(SpreadsheetmlStyles, "patternFill", StylesPruned),
        new(SpreadsheetmlStyles, "patternFill@patternType", StylesPruned),
        new(SpreadsheetmlStyles, "protection", StylesPruned),
        new(SpreadsheetmlStyles, "protection@hidden", StylesPruned),
        new(SpreadsheetmlStyles, "protection@locked", StylesPruned),
        new(SpreadsheetmlStyles, "right", StylesPruned),
        new(SpreadsheetmlStyles, "right@style", StylesPruned),
        new(SpreadsheetmlStyles, "scheme", StylesPruned),
        new(SpreadsheetmlStyles, "scheme@val", StylesPruned),
        new(SpreadsheetmlStyles, "shadow", StylesPruned),
        new(SpreadsheetmlStyles, "shadow@val", StylesPruned),
        new(SpreadsheetmlStyles, "strike", StylesPruned),
        new(SpreadsheetmlStyles, "strike@val", StylesPruned),
        new(SpreadsheetmlStyles, "sz", StylesPruned),
        new(SpreadsheetmlStyles, "sz@val", StylesPruned),
        new(SpreadsheetmlStyles, "top", StylesPruned),
        new(SpreadsheetmlStyles, "top@style", StylesPruned),
        new(SpreadsheetmlStyles, "u", StylesPruned),
        new(SpreadsheetmlStyles, "u@val", StylesPruned),
        new(SpreadsheetmlStyles, "vertAlign", StylesPruned),
        new(SpreadsheetmlStyles, "vertAlign@val", StylesPruned),
        new(SpreadsheetmlWorksheet, "c@s", StyleIndex),
        new(SpreadsheetmlWorksheet, "c@t=inlineStr", InlineString),
        new(SpreadsheetmlWorksheet, "c@t=str", FormulaText),
        new(SpreadsheetmlComments, "comment@shapeId", Deprecated),
        new(SpreadsheetmlWorksheet, "c", NothingToSay),
        new(SpreadsheetmlWorksheet, "c@r", NothingToSay),
        new(SpreadsheetmlWorksheet, "row", NothingToSay),
        new(SpreadsheetmlWorksheet, "row@r", NothingToSay),
        new(CustomProperties, "op:Properties", EmptyPart),
        new(Drawing, "xdr:wsDr", EmptyPart),
        new(PackageRelationships, "Relationship", EmptyPart),
        new(PackageRelationships, "Relationship@Id", EmptyPart),
        new(PackageRelationships, "Relationship@Target", EmptyPart),
        new(PackageRelationships, "Relationship@Type", EmptyPart),
        new(PackageRelationships, "Relationships", EmptyPart),
        new(SpreadsheetmlWorksheet, "drawing", EmptyPart),
        new(SpreadsheetmlWorksheet, "drawing@r:id", EmptyPart),
        new(SpreadsheetmlWorksheet, "col", ColumnsMerged, 68),
        new(SpreadsheetmlWorksheet, "col@bestFit", ColumnsMerged, 34),
        new(SpreadsheetmlWorksheet, "col@customWidth", ColumnsMerged, 70),
        new(SpreadsheetmlWorksheet, "col@max", ColumnsMerged, 68),
        new(SpreadsheetmlWorksheet, "col@min", ColumnsMerged, 68),
        new(SpreadsheetmlWorksheet, "col@style", ColumnsMerged, 32),
        new(SpreadsheetmlWorksheet, "col@width", ColumnsMerged, 68),
        new(SpreadsheetmlWorksheet, "cfRule", FormatsMerged, 10),
        new(SpreadsheetmlWorksheet, "cfRule@dxfId", FormatsMerged, 10),
        new(SpreadsheetmlWorksheet, "cfRule@operator", FormatsMerged, 6),
        new(SpreadsheetmlWorksheet, "cfRule@priority", FormatsMerged, 10),
        new(SpreadsheetmlWorksheet, "cfRule@stopIfTrue", FormatsMerged, 4),
        new(SpreadsheetmlWorksheet, "cfRule@type", FormatsMerged, 10),
        new(SpreadsheetmlWorksheet, "conditionalFormatting", FormatsMerged, 4),
        new(SpreadsheetmlWorksheet, "conditionalFormatting@sqref", FormatsMerged, 4),
        new(SpreadsheetmlWorksheet, "formula", FormatsMerged, 34),
        new(SpreadsheetmlWorksheet, "dataValidation", ValidationsMerged, 326),
        new(SpreadsheetmlWorksheet, "dataValidation@allowBlank", ValidationsMerged, 326),
        new(SpreadsheetmlWorksheet, "dataValidation@error", ValidationsMerged, 40),
        new(SpreadsheetmlWorksheet, "dataValidation@errorTitle", ValidationsMerged, 40),
        new(SpreadsheetmlWorksheet, "dataValidation@operator", ValidationsMerged, 2),
        new(SpreadsheetmlWorksheet, "dataValidation@prompt", ValidationsMerged, 172),
        new(SpreadsheetmlWorksheet, "dataValidation@promptTitle", ValidationsMerged, 40),
        new(SpreadsheetmlWorksheet, "dataValidation@showDropDown", ValidationsMerged, 2),
        new(SpreadsheetmlWorksheet, "dataValidation@showErrorMessage", ValidationsMerged, 182),
        new(SpreadsheetmlWorksheet, "dataValidation@showInputMessage", ValidationsMerged, 326),
        new(SpreadsheetmlWorksheet, "dataValidation@sqref", ValidationsMerged, 326),
        new(SpreadsheetmlWorksheet, "dataValidation@type", ValidationsMerged, 42),
        new(SpreadsheetmlWorksheet, "dataValidations", ValidationsMerged, 4),
        new(SpreadsheetmlWorksheet, "extLst", ValidationsMerged, 4),
        new(SpreadsheetmlWorksheet, "ext[@uri={CCE6A557-97BC-4b89-ADB6-D9C93CAAB3DF}]", ValidationsMerged, 4),
        new(SpreadsheetmlWorksheet, "formula1", ValidationsMerged, 42),
        new(SpreadsheetmlWorksheet, "x14:dataValidation", ValidationsMerged, 10),
        new(SpreadsheetmlWorksheet, "x14:dataValidation@allowBlank", ValidationsMerged, 10),
        new(SpreadsheetmlWorksheet, "x14:dataValidation@showErrorMessage", ValidationsMerged, 10),
        new(SpreadsheetmlWorksheet, "x14:dataValidation@showInputMessage", ValidationsMerged, 10),
        new(SpreadsheetmlWorksheet, "x14:dataValidation@type", ValidationsMerged, 10),
        new(SpreadsheetmlWorksheet, "x14:dataValidations", ValidationsMerged, 4),
        new(SpreadsheetmlWorksheet, "x14:formula1", ValidationsMerged, 10),
        new(SpreadsheetmlWorksheet, "x14:formula2", ValidationsMerged, 2),
        new(SpreadsheetmlWorksheet, "xm:f", ValidationsMerged, 12),
        new(SpreadsheetmlWorksheet, "xm:sqref", ValidationsMerged, 10),
        new(SpreadsheetmlWorksheet, "f@ref", SharedFormulas, 79),
        new(SpreadsheetmlWorksheet, "f@si", SharedFormulas),
        new(SpreadsheetmlWorksheet, "f@t=shared", SharedFormulas),
        new(SpreadsheetmlWorksheet, "f@aca", Volatile),
        new(SpreadsheetmlWorksheet, "f@ca", Volatile),
        new(SpreadsheetmlCalcChain, "c", CalcChain),
        new(SpreadsheetmlCalcChain, "c@a", CalcChain),
        new(SpreadsheetmlCalcChain, "c@i", CalcChain),
        new(SpreadsheetmlCalcChain, "c@l", CalcChain),
        new(SpreadsheetmlCalcChain, "c@r", CalcChain),
        new(SpreadsheetmlCalcChain, "c@s", CalcChain),
        new(SpreadsheetmlWorksheet, "row@spans", Spans),
        new(SpreadsheetmlWorksheet, "row@x14ac:dyDescent", DyDescent),
        new(SpreadsheetmlComments, "t@{http://www.w3.org/XML/1998/namespace}space", PreserveSpace),
        new(SpreadsheetmlSharedStrings, "t@{http://www.w3.org/XML/1998/namespace}space", PreserveSpace),
        new(SpreadsheetmlWorksheet, "f/v@{http://www.w3.org/XML/1998/namespace}space", PreserveSpace),
        new(SpreadsheetmlWorksheet, "oddFooter@{http://www.w3.org/XML/1998/namespace}space", PreserveSpace),
        new(SpreadsheetmlWorksheet, "t@{http://www.w3.org/XML/1998/namespace}space", PreserveSpace),
        new(ExtendedProperties, "ep:Company", EmptyText),
        new(ExtendedProperties, "ep:Manager", EmptyText),
        new(SpreadsheetmlPivotTable, "pivotTableDefinition@missingCaption=", EmptyText),
        new(SpreadsheetmlSheetMain, "definedName[@name=_xlnm.Print_Area]", ReservedName, 4),
        new(SpreadsheetmlSheetMain, "definedName[@name=_xlnm.Print_Titles]", ReservedName, 8),
        new(SpreadsheetmlSheetMain, "definedName[@name=_xlnm.Print_Titles]@localSheetId", ReservedName, 4),
        new(SpreadsheetmlSheetMain, "definedName[@name=_xlnm._FilterDatabase]", ReservedName),
        new(SpreadsheetmlSheetMain, "definedName[@name=_xlnm._FilterDatabase]@function", ReservedName),
        new(SpreadsheetmlSheetMain, "definedName[@name=_xlnm._FilterDatabase]@hidden", ReservedName),
        new(SpreadsheetmlSheetMain, "definedName[@name=_xlnm._FilterDatabase]@localSheetId", ReservedName),
        new(SpreadsheetmlSheetMain, "definedName[@name=_xlnm._FilterDatabase]@vbProcedure", ReservedName),
        new(SpreadsheetmlSheetMain, "sheet@state=hidden", HiddenActiveSheet, 4),
        new(SpreadsheetmlSharedStrings, "si", UnusedString),
        new(SpreadsheetmlSharedStrings, "t", UnusedString),
        new(ContentTypes, "Override", RebuiltContentTypes),
        new(ContentTypes, "Override@ContentType", RebuiltContentTypes),
        new(ContentTypes, "Override@PartName", RebuiltContentTypes),
        new(MsExcelRdrichvaluestructure, "{http://schemas.microsoft.com/office/spreadsheetml/2017/richdata2}k", LegacyCellImage),
        new(MsExcelRdrichvaluestructure, "{http://schemas.microsoft.com/office/spreadsheetml/2017/richdata2}k@n", LegacyCellImage),
        new(MsExcelRdrichvaluestructure, "{http://schemas.microsoft.com/office/spreadsheetml/2017/richdata2}k@t", LegacyCellImage),
        new(MsExcelRdrichvaluestructure, "{http://schemas.microsoft.com/office/spreadsheetml/2017/richdata2}rvStructures", LegacyCellImage),
        new(MsExcelRdrichvaluestructure, "{http://schemas.microsoft.com/office/spreadsheetml/2017/richdata2}s", LegacyCellImage),
        new(MsExcelRdrichvaluestructure, "{http://schemas.microsoft.com/office/spreadsheetml/2017/richdata2}s@t", LegacyCellImage),
        new(MsExcelRichValueRel, "{http://schemas.microsoft.com/office/spreadsheetml/2022/richvaluerel}rel", LegacyCellImage),
        new(MsExcelRichValueRel, "{http://schemas.microsoft.com/office/spreadsheetml/2022/richvaluerel}rel@r:id", LegacyCellImage),
        new(MsExcelRichValueRel, "{http://schemas.microsoft.com/office/spreadsheetml/2022/richvaluerel}richValueRels", LegacyCellImage),
        new(SpreadsheetmlWorksheet, "pageSetup@useFirstPageNumber=true", FirstPageNumber),
        new(SpreadsheetmlWorksheet, "f@dtr", TwoDimensionalTable, 6),
        new(VmlDrawing, "o:lock", CommentShape),
        new(VmlDrawing, "o:lock@selection", CommentShape),
        new(VmlDrawing, "o:lock@v:ext", CommentShape),
        new(VmlDrawing, "v:path", CommentShape),
        new(VmlDrawing, "v:path@gradientshapeok", CommentShape),
        new(VmlDrawing, "v:path@o:connecttype", CommentShape),
        new(VmlDrawing, "v:shape[@type=#_x0000_t202]", CommentShape),
        new(VmlDrawing, "v:shape[@type=#_x0000_t202]@id", CommentShape),
        new(VmlDrawing, "v:shape[@type=#_x0000_t202]@o:spid", CommentShape),
        new(VmlDrawing, "v:shape[@type=#_x0000_t202]@style", CommentShape),
        new(VmlDrawing, "v:shapetype[@id=_x0000_t202]", CommentShape),
        new(VmlDrawing, "v:shapetype[@id=_x0000_t202]@coordsize", CommentShape),
        new(VmlDrawing, "v:shapetype[@id=_x0000_t202]@o:spt", CommentShape),
        new(VmlDrawing, "v:shapetype[@id=_x0000_t202]@path", CommentShape),
        new(VmlDrawing, "v:stroke", CommentShape),
        new(VmlDrawing, "v:stroke@joinstyle", CommentShape),
    ];

    private static readonly KnownLoss[] Known =
    [
        new(SpreadsheetmlWorksheet, "c@t=e", "TBD", LostACellImageIsWrittenAs, 2),
        new(SpreadsheetmlWorksheet, "c@s->cellStyle", "TBD", LostACellSNamedCellStyle, 2829),
        new(SpreadsheetmlWorksheet, "c@s->fill", "TBD", LostASolidFillInTheSystem, 2),
        new(SpreadsheetmlWorksheet, "f/v", "TBD", LostCachedValueOfAFormulaXLibur, 1),
        new(SpreadsheetmlComments, "anchor", "TBD", LostCommentPropertiesCommentPrAndItsAnchor, 4),
        new(SpreadsheetmlComments, "anchor@moveWithCells", "TBD", LostCommentPropertiesCommentPrAndItsAnchor, 2),
        new(SpreadsheetmlComments, "anchor@sizeWithCells", "TBD", LostCommentPropertiesCommentPrAndItsAnchor, 2),
        new(SpreadsheetmlComments, "commentPr", "TBD", LostCommentPropertiesCommentPrAndItsAnchor, 4),
        new(SpreadsheetmlComments, "commentPr@autoFill", "TBD", LostCommentPropertiesCommentPrAndItsAnchor, 4),
        new(SpreadsheetmlComments, "commentPr@autoLine", "TBD", LostCommentPropertiesCommentPrAndItsAnchor, 2),
        new(SpreadsheetmlComments, "commentPr@autoScale", "TBD", LostCommentPropertiesCommentPrAndItsAnchor, 2),
        new(SpreadsheetmlComments, "commentPr@colHidden", "TBD", LostCommentPropertiesCommentPrAndItsAnchor, 2),
        new(SpreadsheetmlComments, "commentPr@defaultSize", "TBD", LostCommentPropertiesCommentPrAndItsAnchor, 2),
        new(SpreadsheetmlComments, "commentPr@locked", "TBD", LostCommentPropertiesCommentPrAndItsAnchor, 2),
        new(SpreadsheetmlComments, "commentPr@rowHidden", "TBD", LostCommentPropertiesCommentPrAndItsAnchor, 2),
        new(SpreadsheetmlComments, "commentPr@textHAlign", "TBD", LostCommentPropertiesCommentPrAndItsAnchor, 2),
        new(SpreadsheetmlComments, "commentPr@textVAlign", "TBD", LostCommentPropertiesCommentPrAndItsAnchor, 2),
        new(SpreadsheetmlComments, "from", "TBD", LostCommentPropertiesCommentPrAndItsAnchor, 2),
        new(SpreadsheetmlComments, "mc:AlternateContent", "TBD", LostCommentPropertiesCommentPrAndItsAnchor, 4),
        new(SpreadsheetmlComments, "mc:Choice", "TBD", LostCommentPropertiesCommentPrAndItsAnchor, 4),
        new(SpreadsheetmlComments, "mc:Choice@Requires", "TBD", LostCommentPropertiesCommentPrAndItsAnchor, 4),
        new(SpreadsheetmlComments, "mc:Fallback", "TBD", LostCommentPropertiesCommentPrAndItsAnchor, 4),
        new(SpreadsheetmlComments, "to", "TBD", LostCommentPropertiesCommentPrAndItsAnchor, 2),
        new(SpreadsheetmlComments, "xdr:col", "TBD", LostCommentPropertiesCommentPrAndItsAnchor, 8),
        new(SpreadsheetmlComments, "xdr:colOff", "TBD", LostCommentPropertiesCommentPrAndItsAnchor, 8),
        new(SpreadsheetmlComments, "xdr:from", "TBD", LostCommentPropertiesCommentPrAndItsAnchor, 2),
        new(SpreadsheetmlComments, "xdr:row", "TBD", LostCommentPropertiesCommentPrAndItsAnchor, 8),
        new(SpreadsheetmlComments, "xdr:rowOff", "TBD", LostCommentPropertiesCommentPrAndItsAnchor, 8),
        new(SpreadsheetmlComments, "xdr:to", "TBD", LostCommentPropertiesCommentPrAndItsAnchor, 2),
        new(VmlDrawing, "div", "TBD", LostCommentTextBoxAlignmentVMLDiv, 70),
        new(VmlDrawing, "div@style", "TBD", LostCommentTextBoxAlignmentVMLDiv, 70),
        new(SpreadsheetmlComments, "charset", "TBD", LostCommentTextRunCharsetAndIndexed, 30),
        new(SpreadsheetmlComments, "charset@val", "TBD", LostCommentTextRunCharsetAndIndexed, 30),
        new(SpreadsheetmlComments, "color", "TBD", LostCommentTextRunCharsetAndIndexed, 48),
        new(SpreadsheetmlComments, "color@indexed", "TBD", LostCommentTextRunCharsetAndIndexed, 48),
        new(SpreadsheetmlWorksheet, "hyperlink@display", "TBD", LostHyperlinkDisplayTextDropped, 4),
        new(SpreadsheetmlPivotCacheDefinition, "d@u", "TBD", LostPivotCacheItemFlagsUContainsNonDate, 3508),
        new(SpreadsheetmlPivotCacheDefinition, "m@u", "TBD", LostPivotCacheItemFlagsUContainsNonDate, 4),
        new(SpreadsheetmlPivotCacheDefinition, "n@u", "TBD", LostPivotCacheItemFlagsUContainsNonDate, 540),
        new(SpreadsheetmlPivotCacheDefinition, "s@u", "TBD", LostPivotCacheItemFlagsUContainsNonDate, 624),
        new(SpreadsheetmlPivotCacheDefinition, "sharedItems@containsNonDate", "TBD", LostPivotCacheItemFlagsUContainsNonDate, 72),
        new(SpreadsheetmlStyles, "alignment@readingOrder", "TBD", LostRebuiltPivotFormatDxfsDropReadingOrder, 194),
        new(SpreadsheetmlStyles, "horizontal", "TBD", LostRebuiltPivotFormatDxfsDropReadingOrder, 72),
        new(SpreadsheetmlStyles, "horizontal@style", "TBD", LostRebuiltPivotFormatDxfsDropReadingOrder, 46),
        new(SpreadsheetmlStyles, "vertical", "TBD", LostRebuiltPivotFormatDxfsDropReadingOrder, 68),
        new(SpreadsheetmlStyles, "vertical@style", "TBD", LostRebuiltPivotFormatDxfsDropReadingOrder, 46),
        new(SpreadsheetmlSharedStrings, "charset", "TBD", LostRichTextRunCharsetDropped, 6),
        new(SpreadsheetmlSharedStrings, "charset@val", "TBD", LostRichTextRunCharsetDropped, 6),
        new(MsExcelRdrichvalue, "{http://schemas.microsoft.com/office/spreadsheetml/2017/richdata}rv", "TBD", LostRichValueDetailOfAVALUE, 14),
        new(MsExcelRdrichvalue, "{http://schemas.microsoft.com/office/spreadsheetml/2017/richdata}rv@s", "TBD", LostRichValueDetailOfAVALUE, 14),
        new(MsExcelRdrichvalue, "{http://schemas.microsoft.com/office/spreadsheetml/2017/richdata}v", "TBD", LostRichValueDetailOfAVALUE, 50),
        new(MsExcelRdrichvaluestructure, "{http://schemas.microsoft.com/office/spreadsheetml/2017/richdata}k", "TBD", LostRichValueDetailOfAVALUE, 24),
        new(MsExcelRdrichvaluestructure, "{http://schemas.microsoft.com/office/spreadsheetml/2017/richdata}k@n", "TBD", LostRichValueDetailOfAVALUE, 24),
        new(MsExcelRdrichvaluestructure, "{http://schemas.microsoft.com/office/spreadsheetml/2017/richdata}k@t", "TBD", LostRichValueDetailOfAVALUE, 24),
        new(MsExcelRdrichvaluestructure, "{http://schemas.microsoft.com/office/spreadsheetml/2017/richdata}s", "TBD", LostRichValueDetailOfAVALUE, 4),
        new(MsExcelRdrichvaluestructure, "{http://schemas.microsoft.com/office/spreadsheetml/2017/richdata}s@t", "TBD", LostRichValueDetailOfAVALUE, 4),
        new(SpreadsheetmlSheetMetadata, "bk", "TBD", LostRichValueDetailOfAVALUE, 14),
        new(SpreadsheetmlSheetMetadata, "ext", "TBD", LostRichValueDetailOfAVALUE, 14),
        new(SpreadsheetmlSheetMetadata, "ext@uri", "TBD", LostRichValueDetailOfAVALUE, 14),
        new(SpreadsheetmlSheetMetadata, "extLst", "TBD", LostRichValueDetailOfAVALUE, 14),
        new(SpreadsheetmlSheetMetadata, "{http://schemas.microsoft.com/office/spreadsheetml/2017/richdata}rvb", "TBD", LostRichValueDetailOfAVALUE, 14),
        new(SpreadsheetmlSheetMetadata, "{http://schemas.microsoft.com/office/spreadsheetml/2017/richdata}rvb@i", "TBD", LostRichValueDetailOfAVALUE, 14),
        new(SpreadsheetmlWorksheet, "c@vm", "TBD", LostRichValueDetailOfAVALUE, 16),
        new(SpreadsheetmlWorksheet, "row@customFormat", "TBD", LostRowStyleCustomFormatSDroppedWhen, 100),
        new(SpreadsheetmlWorksheet, "row@s", "TBD", LostRowStyleCustomFormatSDroppedWhen, 94),
        new(SpreadsheetmlWorksheet, "row@thickBot", "TBD", LostRowThickBotAndThickTopDroppedWhen, 108),
        new(SpreadsheetmlWorksheet, "row@thickTop", "TBD", LostRowThickBotAndThickTopDroppedWhen, 30),
        new(SpreadsheetmlWorksheet, "selection@pane", "TBD", LostSelectionOfAFrozenPaneDropped, 4),
        new(SpreadsheetmlTable, "tableColumn@dataDxfId", "TBD", LostTableColumnDataFormatDataDxfIdDropped, 8),
        new(SpreadsheetmlWorksheet, "c@s->font", "TBD", LostTheFontFamilyAndInOne, 1703),
    ];

    internal static IEnumerable<(string ContentType, string Key)> Keys =>
        Noise.Select(n => (n.ContentType, n.Key)).Concat(Known.Select(k => (k.ContentType, k.Key)));

    internal static IReadOnlySet<(string ContentType, string Key)> Qualified { get; } =
        Keys.Where(k => k.Key.Contains('=')).ToHashSet();

    private static readonly HashSet<(string, string)> All = Keys.ToHashSet();

    /// <summary>
    /// The most of each bounded key the whole corpus may lose, summed over every fixture and both saves.
    /// </summary>
    internal static IReadOnlyDictionary<(string ContentType, string Key), int> Bounds { get; } =
        Noise.Where(n => n.Bound is not null).Select(n => (n.ContentType, n.Key, Bound: n.Bound!.Value))
            .Concat(Known.Select(k => (k.ContentType, k.Key, k.Bound)))
            .ToDictionary(b => (b.ContentType, b.Key), b => b.Bound);

    internal static bool Explains(string contentType, string key) => All.Contains((contentType, Normalise(key)));

    internal static string Normalise(string key) =>
        key.StartsWith(RoundTripCensus.BinaryKeyPrefix, StringComparison.Ordinal) ? "binary" : key;
}

/// <summary>A loss that is not one. A key that still means something to a user also carries a bound.</summary>
internal sealed record AllowedLoss(string ContentType, string Key, string Reason, int? Bound = null);

/// <summary>A real loss with an issue that tracks it, bounded so the gap cannot grow unseen.</summary>
internal sealed record KnownLoss(string ContentType, string Key, string Issue, string Gap, int Bound);
