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
    private const string InvalidColor = "an invalid colour, <color rgb=\"0\"/> in InvalidColors.xlsx, is written as the colour it decodes to, so the cell's font no longer matches the loaded one";
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
    private const string ConditionalAlignment = "a conditional format cannot apply alignment, so its dxf is written without one";
    private const string SharedFormulas = "a shared formula is written as a formula in each cell";
    private const string DirtyFormula = "a formula an edit left dirty, which XLibur cannot evaluate, is written with no cached value and no t, so Excel recalculates it (ADR 0001, #962)";
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
    private const string LostASolidFillInTheSystem = "a solid fill in the system background colour (indexed 65) written as no fill when the sheet is written from the model";
    private const string LostCommentPropertiesCommentPrAndItsAnchor = "comment properties (commentPr and its anchor) dropped";
    private const string LostCommentTextBoxAlignmentVMLDiv = "comment text box alignment (VML div style) dropped";
    private const string LostCommentTextRunCharsetAndIndexed = "comment text run charset and indexed colour dropped";
    private const string LostPivotCacheItemFlagsUContainsNonDate = "pivot cache item flags (u, containsNonDate) dropped";
    private const string LostRebuiltPivotFormatDxfsDropReadingOrder = "rebuilt pivot format dxfs drop readingOrder and the vertical and horizontal border sides";
    private const string LostRichTextRunCharsetDropped = "rich text run charset dropped";
    private const string LostRichValueDetailOfAVALUE = "rich-value detail of a #VALUE! error (subType, argument) dropped";
    private const string LostRowStyleCustomFormatSDroppedWhen = "row style (customFormat, s) dropped when the sheet is written from the model";
    private const string LostRowThickBotAndThickTopDroppedWhen = "row thickBot and thickTop dropped when the sheet is written from the model";
    private const string LostTableColumnDataFormatDataDxfIdDropped = "table column data format (dataDxfId) dropped";

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
        new(SpreadsheetmlWorksheet, "c@s->font", InvalidColor, 1),
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
        new(SpreadsheetmlStyles, "dxfs@count", Count, 380),
        new(SpreadsheetmlWorksheet, "dataValidations@count", Count),
        new(SpreadsheetmlWorksheet, "x14:dataValidations@count", Count, 4),
        new(SpreadsheetmlStyles, "alignment", StylesPruned, 192),
        new(SpreadsheetmlStyles, "alignment@horizontal", StylesPruned, 138),
        new(SpreadsheetmlStyles, "alignment@indent", StylesPruned, 82),
        new(SpreadsheetmlStyles, "alignment@justifyLastLine", StylesPruned, 82),
        new(SpreadsheetmlStyles, "alignment@shrinkToFit", StylesPruned, 82),
        new(SpreadsheetmlStyles, "alignment@textRotation", StylesPruned, 82),
        new(SpreadsheetmlStyles, "alignment@vertical", StylesPruned, 82),
        new(SpreadsheetmlStyles, "alignment@wrapText", StylesPruned, 134),
        new(SpreadsheetmlStyles, "b", StylesPruned, 22),
        new(SpreadsheetmlStyles, "b@val", StylesPruned, 2),
        new(SpreadsheetmlStyles, "bgColor", StylesPruned, 226),
        new(SpreadsheetmlStyles, "bgColor@auto", StylesPruned, 10),
        new(SpreadsheetmlStyles, "bgColor@indexed", StylesPruned, 86),
        new(SpreadsheetmlStyles, "bgColor@rgb", StylesPruned, 116),
        new(SpreadsheetmlStyles, "bgColor@theme", StylesPruned, 14),
        new(SpreadsheetmlStyles, "bgColor@tint", StylesPruned, 6),
        new(SpreadsheetmlStyles, "border", StylesPruned, 84),
        new(SpreadsheetmlStyles, "border@diagonalDown", StylesPruned, 6),
        new(SpreadsheetmlStyles, "border@diagonalUp", StylesPruned, 6),
        new(SpreadsheetmlStyles, "bottom", StylesPruned, 84),
        new(SpreadsheetmlStyles, "bottom@style", StylesPruned, 50),
        new(SpreadsheetmlStyles, "charset", StylesPruned, 6),
        new(SpreadsheetmlStyles, "charset@val", StylesPruned, 6),
        new(SpreadsheetmlStyles, "color", StylesPruned, 256),
        new(SpreadsheetmlStyles, "color@auto", StylesPruned, 32),
        new(SpreadsheetmlStyles, "color@indexed", StylesPruned, 212),
        new(SpreadsheetmlStyles, "color@theme", StylesPruned, 26),
        new(SpreadsheetmlStyles, "color@tint", StylesPruned, 8),
        new(SpreadsheetmlStyles, "condense", StylesPruned, 6),
        new(SpreadsheetmlStyles, "condense@val", StylesPruned, 6),
        new(SpreadsheetmlStyles, "diagonal", StylesPruned, 8),
        new(SpreadsheetmlStyles, "diagonal@style", StylesPruned, 8),
        new(SpreadsheetmlStyles, "dxf", StylesPruned, 1186),
        new(SpreadsheetmlStyles, "dxfs", StylesPruned, 380),
        new(SpreadsheetmlStyles, "extend", StylesPruned, 6),
        new(SpreadsheetmlStyles, "extend@val", StylesPruned, 6),
        new(SpreadsheetmlStyles, "family", StylesPruned, 4),
        new(SpreadsheetmlStyles, "family@val", StylesPruned, 4),
        new(SpreadsheetmlStyles, "fgColor", StylesPruned, 92),
        new(SpreadsheetmlStyles, "fgColor@indexed", StylesPruned, 86),
        new(SpreadsheetmlStyles, "fgColor@rgb", StylesPruned, 2),
        new(SpreadsheetmlStyles, "fgColor@theme", StylesPruned, 4),
        new(SpreadsheetmlStyles, "fgColor@tint", StylesPruned, 2),
        new(SpreadsheetmlStyles, "fill", StylesPruned, 230),
        new(SpreadsheetmlStyles, "font", StylesPruned, 76),
        new(SpreadsheetmlStyles, "i", StylesPruned, 22),
        new(SpreadsheetmlStyles, "i@val", StylesPruned, 8),
        new(SpreadsheetmlStyles, "left", StylesPruned, 68),
        new(SpreadsheetmlStyles, "left@style", StylesPruned, 44),
        new(SpreadsheetmlStyles, "name", StylesPruned, 10),
        new(SpreadsheetmlStyles, "name@val", StylesPruned, 10),
        new(SpreadsheetmlStyles, "numFmt", StylesPruned, 794),
        new(SpreadsheetmlStyles, "numFmt@formatCode", StylesPruned, 794),
        new(SpreadsheetmlStyles, "numFmt@numFmtId", StylesPruned, 794),
        new(SpreadsheetmlStyles, "outline", StylesPruned, 6),
        new(SpreadsheetmlStyles, "outline@val", StylesPruned, 6),
        new(SpreadsheetmlStyles, "patternFill", StylesPruned, 230),
        new(SpreadsheetmlStyles, "patternFill@patternType", StylesPruned, 154),
        new(SpreadsheetmlStyles, "protection", StylesPruned, 4),
        new(SpreadsheetmlStyles, "protection@hidden", StylesPruned, 4),
        new(SpreadsheetmlStyles, "protection@locked", StylesPruned, 4),
        new(SpreadsheetmlStyles, "right", StylesPruned, 84),
        new(SpreadsheetmlStyles, "right@style", StylesPruned, 44),
        new(SpreadsheetmlStyles, "scheme", StylesPruned, 14),
        new(SpreadsheetmlStyles, "scheme@val", StylesPruned, 14),
        new(SpreadsheetmlStyles, "shadow", StylesPruned, 6),
        new(SpreadsheetmlStyles, "shadow@val", StylesPruned, 6),
        new(SpreadsheetmlStyles, "strike", StylesPruned, 14),
        new(SpreadsheetmlStyles, "strike@val", StylesPruned, 14),
        new(SpreadsheetmlStyles, "sz", StylesPruned, 32),
        new(SpreadsheetmlStyles, "sz@val", StylesPruned, 32),
        new(SpreadsheetmlStyles, "top", StylesPruned, 84),
        new(SpreadsheetmlStyles, "top@style", StylesPruned, 44),
        new(SpreadsheetmlStyles, "u", StylesPruned, 14),
        new(SpreadsheetmlStyles, "u@val", StylesPruned, 14),
        new(SpreadsheetmlStyles, "vertAlign", StylesPruned, 10),
        new(SpreadsheetmlStyles, "vertAlign@val", StylesPruned, 10),
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
        new(SpreadsheetmlWorksheet, "col", ColumnsMerged, 70),
        new(SpreadsheetmlWorksheet, "col@bestFit", ColumnsMerged, 34),
        new(SpreadsheetmlWorksheet, "col@customWidth", ColumnsMerged, 70),
        new(SpreadsheetmlWorksheet, "col@max", ColumnsMerged, 70),
        new(SpreadsheetmlWorksheet, "col@min", ColumnsMerged, 70),
        new(SpreadsheetmlWorksheet, "col@style", ColumnsMerged, 32),
        new(SpreadsheetmlWorksheet, "col@width", ColumnsMerged, 70),
        new(SpreadsheetmlWorksheet, "cfRule", FormatsMerged, 10),
        new(SpreadsheetmlWorksheet, "cfRule@dxfId", FormatsMerged, 10),
        new(SpreadsheetmlWorksheet, "cfRule@dxfId->fill", FormatsMerged, 10),
        new(SpreadsheetmlWorksheet, "cfRule@dxfId->font", FormatsMerged, 6),
        new(SpreadsheetmlWorksheet, "cfRule@dxfId->numFmt", FormatsMerged, 6),
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
        new(SpreadsheetmlWorksheet, "cfRule@dxfId->alignment", ConditionalAlignment, 1300),
        new(SpreadsheetmlWorksheet, "f@ref", SharedFormulas, 79),
        new(SpreadsheetmlWorksheet, "f@si", SharedFormulas),
        new(SpreadsheetmlWorksheet, "f@t=shared", SharedFormulas),
        new(SpreadsheetmlWorksheet, "f@aca", Volatile),
        new(SpreadsheetmlWorksheet, "f@ca", Volatile),
        new(SpreadsheetmlWorksheet, "f/v", DirtyFormula, 1),
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
        // Bound 3: two are #960; the third is the t="e" of the edited bug-fixes.xlsx's averageif cell,
        // which goes with its cached value (DirtyFormula, f/v). Lower it to 1 when #960 is fixed.
        new(SpreadsheetmlWorksheet, "c@t=e", "#960", LostACellImageIsWrittenAs, 3),
        new(SpreadsheetmlWorksheet, "c@s->fill", "#957", LostASolidFillInTheSystem, 2),
        new(SpreadsheetmlComments, "anchor", "#953", LostCommentPropertiesCommentPrAndItsAnchor, 4),
        new(SpreadsheetmlComments, "anchor@moveWithCells", "#953", LostCommentPropertiesCommentPrAndItsAnchor, 2),
        new(SpreadsheetmlComments, "anchor@sizeWithCells", "#953", LostCommentPropertiesCommentPrAndItsAnchor, 2),
        new(SpreadsheetmlComments, "commentPr", "#953", LostCommentPropertiesCommentPrAndItsAnchor, 4),
        new(SpreadsheetmlComments, "commentPr@autoFill", "#953", LostCommentPropertiesCommentPrAndItsAnchor, 4),
        new(SpreadsheetmlComments, "commentPr@autoLine", "#953", LostCommentPropertiesCommentPrAndItsAnchor, 2),
        new(SpreadsheetmlComments, "commentPr@autoScale", "#953", LostCommentPropertiesCommentPrAndItsAnchor, 2),
        new(SpreadsheetmlComments, "commentPr@colHidden", "#953", LostCommentPropertiesCommentPrAndItsAnchor, 2),
        new(SpreadsheetmlComments, "commentPr@defaultSize", "#953", LostCommentPropertiesCommentPrAndItsAnchor, 2),
        new(SpreadsheetmlComments, "commentPr@locked", "#953", LostCommentPropertiesCommentPrAndItsAnchor, 2),
        new(SpreadsheetmlComments, "commentPr@rowHidden", "#953", LostCommentPropertiesCommentPrAndItsAnchor, 2),
        new(SpreadsheetmlComments, "commentPr@textHAlign", "#953", LostCommentPropertiesCommentPrAndItsAnchor, 2),
        new(SpreadsheetmlComments, "commentPr@textVAlign", "#953", LostCommentPropertiesCommentPrAndItsAnchor, 2),
        new(SpreadsheetmlComments, "from", "#953", LostCommentPropertiesCommentPrAndItsAnchor, 2),
        new(SpreadsheetmlComments, "mc:AlternateContent", "#953", LostCommentPropertiesCommentPrAndItsAnchor, 4),
        new(SpreadsheetmlComments, "mc:Choice", "#953", LostCommentPropertiesCommentPrAndItsAnchor, 4),
        new(SpreadsheetmlComments, "mc:Choice@Requires", "#953", LostCommentPropertiesCommentPrAndItsAnchor, 4),
        new(SpreadsheetmlComments, "mc:Fallback", "#953", LostCommentPropertiesCommentPrAndItsAnchor, 4),
        new(SpreadsheetmlComments, "to", "#953", LostCommentPropertiesCommentPrAndItsAnchor, 2),
        new(SpreadsheetmlComments, "xdr:col", "#953", LostCommentPropertiesCommentPrAndItsAnchor, 8),
        new(SpreadsheetmlComments, "xdr:colOff", "#953", LostCommentPropertiesCommentPrAndItsAnchor, 8),
        new(SpreadsheetmlComments, "xdr:from", "#953", LostCommentPropertiesCommentPrAndItsAnchor, 2),
        new(SpreadsheetmlComments, "xdr:row", "#953", LostCommentPropertiesCommentPrAndItsAnchor, 8),
        new(SpreadsheetmlComments, "xdr:rowOff", "#953", LostCommentPropertiesCommentPrAndItsAnchor, 8),
        new(SpreadsheetmlComments, "xdr:to", "#953", LostCommentPropertiesCommentPrAndItsAnchor, 2),
        new(VmlDrawing, "div", "#953", LostCommentTextBoxAlignmentVMLDiv, 70),
        new(VmlDrawing, "div@style", "#953", LostCommentTextBoxAlignmentVMLDiv, 70),
        new(SpreadsheetmlComments, "charset", "#957", LostCommentTextRunCharsetAndIndexed, 30),
        new(SpreadsheetmlComments, "charset@val", "#957", LostCommentTextRunCharsetAndIndexed, 30),
        new(SpreadsheetmlComments, "color", "#957", LostCommentTextRunCharsetAndIndexed, 48),
        new(SpreadsheetmlComments, "color@indexed", "#957", LostCommentTextRunCharsetAndIndexed, 48),
        new(SpreadsheetmlPivotCacheDefinition, "d@u", "#955", LostPivotCacheItemFlagsUContainsNonDate, 3508),
        new(SpreadsheetmlPivotCacheDefinition, "m@u", "#955", LostPivotCacheItemFlagsUContainsNonDate, 4),
        new(SpreadsheetmlPivotCacheDefinition, "n@u", "#955", LostPivotCacheItemFlagsUContainsNonDate, 540),
        new(SpreadsheetmlPivotCacheDefinition, "s@u", "#955", LostPivotCacheItemFlagsUContainsNonDate, 624),
        new(SpreadsheetmlPivotCacheDefinition, "sharedItems@containsNonDate", "#955", LostPivotCacheItemFlagsUContainsNonDate, 72),
        new(SpreadsheetmlPivotTable, "format@dxfId->alignment", "#956", LostRebuiltPivotFormatDxfsDropReadingOrder, 32),
        new(SpreadsheetmlPivotTable, "format@dxfId->border", "#956", LostRebuiltPivotFormatDxfsDropReadingOrder, 14),
        new(SpreadsheetmlStyles, "alignment@readingOrder", "#956", LostRebuiltPivotFormatDxfsDropReadingOrder, 196),
        new(SpreadsheetmlStyles, "horizontal", "#956", LostRebuiltPivotFormatDxfsDropReadingOrder, 70),
        new(SpreadsheetmlStyles, "horizontal@style", "#956", LostRebuiltPivotFormatDxfsDropReadingOrder, 46),
        new(SpreadsheetmlStyles, "vertical", "#956", LostRebuiltPivotFormatDxfsDropReadingOrder, 66),
        new(SpreadsheetmlStyles, "vertical@style", "#956", LostRebuiltPivotFormatDxfsDropReadingOrder, 46),
        new(SpreadsheetmlSharedStrings, "charset", "#957", LostRichTextRunCharsetDropped, 6),
        new(SpreadsheetmlSharedStrings, "charset@val", "#957", LostRichTextRunCharsetDropped, 6),
        new(MsExcelRdrichvalue, "{http://schemas.microsoft.com/office/spreadsheetml/2017/richdata}rv", "#954", LostRichValueDetailOfAVALUE, 14),
        new(MsExcelRdrichvalue, "{http://schemas.microsoft.com/office/spreadsheetml/2017/richdata}rv@s", "#954", LostRichValueDetailOfAVALUE, 14),
        new(MsExcelRdrichvalue, "{http://schemas.microsoft.com/office/spreadsheetml/2017/richdata}v", "#954", LostRichValueDetailOfAVALUE, 50),
        new(MsExcelRdrichvaluestructure, "{http://schemas.microsoft.com/office/spreadsheetml/2017/richdata}k", "#954", LostRichValueDetailOfAVALUE, 24),
        new(MsExcelRdrichvaluestructure, "{http://schemas.microsoft.com/office/spreadsheetml/2017/richdata}k@n", "#954", LostRichValueDetailOfAVALUE, 24),
        new(MsExcelRdrichvaluestructure, "{http://schemas.microsoft.com/office/spreadsheetml/2017/richdata}k@t", "#954", LostRichValueDetailOfAVALUE, 24),
        new(MsExcelRdrichvaluestructure, "{http://schemas.microsoft.com/office/spreadsheetml/2017/richdata}s", "#954", LostRichValueDetailOfAVALUE, 4),
        new(MsExcelRdrichvaluestructure, "{http://schemas.microsoft.com/office/spreadsheetml/2017/richdata}s@t", "#954", LostRichValueDetailOfAVALUE, 4),
        new(SpreadsheetmlSheetMetadata, "bk", "#954", LostRichValueDetailOfAVALUE, 14),
        new(SpreadsheetmlSheetMetadata, "ext", "#954", LostRichValueDetailOfAVALUE, 14),
        new(SpreadsheetmlSheetMetadata, "ext@uri", "#954", LostRichValueDetailOfAVALUE, 14),
        new(SpreadsheetmlSheetMetadata, "extLst", "#954", LostRichValueDetailOfAVALUE, 14),
        new(SpreadsheetmlSheetMetadata, "{http://schemas.microsoft.com/office/spreadsheetml/2017/richdata}rvb", "#954", LostRichValueDetailOfAVALUE, 14),
        new(SpreadsheetmlSheetMetadata, "{http://schemas.microsoft.com/office/spreadsheetml/2017/richdata}rvb@i", "#954", LostRichValueDetailOfAVALUE, 14),
        new(SpreadsheetmlWorksheet, "c@vm", "#954", LostRichValueDetailOfAVALUE, 16),
        new(SpreadsheetmlWorksheet, "row@customFormat", "#958", LostRowStyleCustomFormatSDroppedWhen, 100),
        new(SpreadsheetmlWorksheet, "row@s", "#958", LostRowStyleCustomFormatSDroppedWhen, 94),
        new(SpreadsheetmlWorksheet, "row@thickBot", "#959", LostRowThickBotAndThickTopDroppedWhen, 108),
        new(SpreadsheetmlWorksheet, "row@thickTop", "#959", LostRowThickBotAndThickTopDroppedWhen, 30),
        new(SpreadsheetmlTable, "tableColumn@dataDxfId", "#961", LostTableColumnDataFormatDataDxfIdDropped, 8),
        new(SpreadsheetmlTable, "tableColumn@dataDxfId->alignment", "#961", LostTableColumnDataFormatDataDxfIdDropped, 98),
        new(SpreadsheetmlTable, "tableColumn@dataDxfId->border", "#961", LostTableColumnDataFormatDataDxfIdDropped, 14),
        new(SpreadsheetmlTable, "tableColumn@dataDxfId->fill", "#961", LostTableColumnDataFormatDataDxfIdDropped, 90),
        new(SpreadsheetmlTable, "tableColumn@dataDxfId->font", "#961", LostTableColumnDataFormatDataDxfIdDropped, 14),
        new(SpreadsheetmlTable, "tableColumn@dataDxfId->numFmt", "#961", LostTableColumnDataFormatDataDxfIdDropped, 56),
        new(SpreadsheetmlTable, "tableColumn@dataDxfId->protection", "#961", LostTableColumnDataFormatDataDxfIdDropped, 86),
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
