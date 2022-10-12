using Google.Apis.Sheets.v4.Data;
using Google.Apis.Sheets.v4;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Runtime.InteropServices;
using static GoogleApis.SheetsRange;

namespace GoogleApis
{
    [Guid("8584BC78-F81F-4754-B353-86E6E501CFCB")]
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    public interface ISheetsRange
    {
        string Name { get; set; }
        GSRangeType Type { get; set; }
        int SheetId { get; set; }
        string A1NotationAddress { get; set; }
        string R1C1NotationAddress { get; set; }
        object[,] Values { get; set; }
        object ActiveService { get; set; }
        GoogleSheetsFile Parent { get; set; }
        int ColumnsCount { get; set; }
        int RowsCount { get; set; }
        void SetValues(object RangeValues);
        void SetFormats(object RGB, string FontFamily = null, int FontSize = 0, bool Bold = false, bool Italic = false, bool Underline = false, bool Strikethrough = false, int TextAngle = 0, bool TextVertical = false, GSTextDirection TextDirection = GSTextDirection.gsTextDirectionUnspecified);
        object[,] GetValues();
        object[,] ConvertToZeroBased2DArray(object ExcelRangeValue);
        void AddFilterView(object HiddenValues = null, int FilterColumnIndex = 1, GSConditionType ConditionType = GSConditionType.gsConditionTypeUnspecified, object ConditionValues = null);
        void Sort(GSSortOrder SortOrder, int ColumnIndex);
        void DeleteFilterView(string FilterViewName);
        void AddConditionalFormatRule(string RuleName, object ConditionValues, object RGB, int Position = 0, GSConditionType ConditionType = GSConditionType.gsConditionTypeUnspecified);
        void FindReplace(string Find, string Replacement, bool SearchByRegex = false, bool IncludeFormulas = false, bool MatchCase = false, bool MatchEntireCell = false);
    }
    [Guid("6C84F220-EACB-44E3-BAF9-1A4221D725FE")]
    [ClassInterface(ClassInterfaceType.None)]
    public class SheetsRange : ISheetsRange
    {
        public string Name { get; set; }
        public GSRangeType Type { get; set; }
        public int SheetId { get; set; }
        public string A1NotationAddress { get; set; }
        public string R1C1NotationAddress { get; set; }
        public object[,] Values { get; set; }
        public object ActiveService { get; set; }
        public GoogleSheetsFile Parent { get; set; }
        public int ColumnsCount { get; set; }
        public int RowsCount { get; set; }
        internal int StartColumnIndex { get; set; }
        internal int StartRowIndex { get; set; }
        internal int EndColumnIndex { get; set; }
        internal int EndRowIndex { get; set; }
        public enum GSSortOrder
        {
            gsSortOrderAscending,
            gsSortOrderDescending
        }
        public enum GSConditionType
        {
            gsConditionTypeUnspecified,
            gsConditionTypeNumberGreater,
            gsConditionTypeNumberGreaterThanEQ,
            gsConditionTypeNumberLess,
            gsConditionTypeNumberLessThanEQ,
            gsConditionTypeNumberEQ,
            gsConditionTypeNumberNotEQ,
            gsConditionTypeNumberBetween,
            gsConditionTypeNumberNotBetween,
            gsConditionTypeTextContains,
            gsConditionTypeTextNotContains,
            gsConditionTypeTextStartsWith,
            gsConditionTypeTextEndsWith,
            gsConditionTypeTextEQ,
            gsConditionTypeTextIsEmail,
            gsConditionTypeTextIsURL,
            gsConditionTypeDateEQ,
            gsConditionTypeDateBefore,
            gsConditionTypeDateAfter,
            gsConditionTypeDateOnOrBefore,
            gsConditionTypeDateOnOrAfter,
            gsConditionTypeDateBetween,
            gsConditionTypeDateNotBetween,
            gsConditionTypeDateIsValid,
            gsConditionTypeOneOfRange,
            gsConditionTypeOneOfList,
            gsConditionTypeBlank,
            gsConditionTypeNotBlank,
            gsConditionTypeCustomFormula,
            gsConditionTypeBoolean,
            gsConditionTypeTextNotEQ,
            gsConditionTypeDateNotEQ
        }
        public enum GSTextDirection
        {
            gsTextDirectionUnspecified,
            gsTextDirectionLeftToRight,
            gsTextDirectionRightToLeft
        }
        public enum GSRangeType
        {
            gsRangeTypeNormal,
            gsRangeTypeNamedRange,
            gsRangeTypeProtectedRange
        }
        public void SetValues(object RangeValues)
        {
            Type ValuesType = RangeValues.GetType();
            if (ValuesType.IsArray)
            {
                object[,] Values2d = (object[,])RangeValues;
                Google.Apis.Sheets.v4.SheetsService objSheetsService = (Google.Apis.Sheets.v4.SheetsService)ActiveService;
                Google.Apis.Sheets.v4.Data.ValueRange objValueRange = new ValueRange();
                IList<IList<object>> ListsValues = new List<IList<object>>(Values2d.GetLength(0));
                for (int i = 0; i < Values2d.GetLength(0); i++)
                {
                    List<object> ListObject = new List<object>(Values2d.GetLength(1));
                    ListsValues.Add(ListObject);
                    for (int j = 0; j < Values2d.GetLength(1); j++)
                    {
                        ListObject.Add(Values2d[i, j]);
                    }
                }
                objValueRange.Values = ListsValues;
                SpreadsheetsResource.ValuesResource.UpdateRequest objUpdateRequest = objSheetsService.Spreadsheets.Values.Update(objValueRange, Parent.Id, A1NotationAddress);
                objUpdateRequest.ValueInputOption = SpreadsheetsResource.ValuesResource.UpdateRequest.ValueInputOptionEnum.USERENTERED;
                objUpdateRequest.Execute();
                Values = Values2d;
            }
            else
            {
                throw new ArgumentException("Invalid argument data type. Accepted data type: Zero-based two dimension array of variant.");
            }
        }
        public object[,] ConvertToZeroBased2DArray(object ExcelRangeValue)
        {
            Type ExcelRangeValueType = ExcelRangeValue.GetType();
            if (ExcelRangeValueType.IsArray)
            {
                object[,] arrTemp = (object[,])ExcelRangeValue;
                int intRowsCount = arrTemp.GetLength(0) - 1;
                int intColumnsCount = arrTemp.GetLength(1) - 1;
                object[,] arrResult = new object[intRowsCount + 1, intColumnsCount + 1];
                for (int i = 0; i < intRowsCount + 1; i++)
                {
                    for (int j = 0; j < intColumnsCount + 1; j++)
                    {
                        arrResult[i, j] = arrTemp[i + 1, j + 1];
                    }
                }
                return arrResult;
            }
            else { return null; }
        }
        public void SetFormats(object RGB, string FontFamily = null, int FontSize = 0, bool Bold = false, bool Italic = false, bool Underline = false , bool Strikethrough = false, int TextAngle = 0, bool TextVertical = false, GSTextDirection TextDirection = GSTextDirection.gsTextDirectionUnspecified)
        {
            Google.Apis.Sheets.v4.SheetsService objSheetsService = (Google.Apis.Sheets.v4.SheetsService)ActiveService;
            TextFormat objTextFormat = new TextFormat();
            objTextFormat.Bold = Bold;
            if (FontFamily != null) { objTextFormat.FontFamily = FontFamily; }
            if (FontSize > 0) { objTextFormat.FontSize = FontSize; }
            objTextFormat.Italic = Italic;
            objTextFormat.Underline = Underline;
            objTextFormat.Strikethrough = Strikethrough;
            TextRotation objTextRotation = new TextRotation
            {
                Angle = TextAngle,
                Vertical = TextVertical
            };
            Google.Apis.Sheets.v4.Data.CellFormat objCellFormat = new CellFormat
            {
                TextFormat = objTextFormat,
                TextRotation = objTextRotation
            };
            switch (TextDirection)
            {
                case GSTextDirection.gsTextDirectionUnspecified:
                    objCellFormat.TextDirection = "TEXT_DIRECTION_UNSPECIFIED";
                    break;
                case GSTextDirection.gsTextDirectionLeftToRight:
                    objCellFormat.TextDirection = "LEFT_TO_RIGHT";
                    break;
                case GSTextDirection.gsTextDirectionRightToLeft:
                    objCellFormat.TextDirection = "RIGHT_TO_LEFT";
                    break;
            }
            Type RGBType = RGB.GetType();
            Color objColor = new Color();
            if (RGBType.IsArray)
            {
                float[] arrRGB = (float[])RGB;
                objColor.Red = arrRGB[0] / 255;
                objColor.Blue = arrRGB[1] / 255;
                objColor.Green = arrRGB[2] / 255;
                objColor.Alpha = (float)0.5;
                objTextFormat.ForegroundColorStyle = new ColorStyle
                {
                    RgbColor = objColor
                };
            }
            else { throw new ArgumentException("RGB: Wrong argument type. Accepted data type: Zero-based one dimension array of variant."); }
            Google.Apis.Sheets.v4.Data.Request objRequest = new Request
            {
                RepeatCell = new RepeatCellRequest
                {
                    Range = new GridRange
                    {
                        SheetId = SheetId,
                        StartColumnIndex = StartColumnIndex - 1,
                        EndColumnIndex = EndColumnIndex,
                        StartRowIndex = StartRowIndex - 1,
                        EndRowIndex = EndRowIndex
                    },
                    Cell = new CellData
                    {
                        UserEnteredFormat = objCellFormat
                    },
                    Fields = "UserEnteredFormat(TextFormat)"
                },
            };
            Google.Apis.Sheets.v4.Data.BatchUpdateSpreadsheetRequest objBatchUpdateSpreadsheetRequest = new BatchUpdateSpreadsheetRequest
            {
                Requests = new List<Request> { objRequest}
            };
            Google.Apis.Sheets.v4.SpreadsheetsResource.BatchUpdateRequest objBatchUpdate = objSheetsService.Spreadsheets.BatchUpdate(objBatchUpdateSpreadsheetRequest, Parent.Id);
            objBatchUpdate.Execute();
        }
        public object[,] GetValues()
        {
            try
            {
                Google.Apis.Sheets.v4.SheetsService objSheetsService = (Google.Apis.Sheets.v4.SheetsService)ActiveService;
                Google.Apis.Sheets.v4.SpreadsheetsResource.GetRequest objGetSpreadsheetRequest = objSheetsService.Spreadsheets.Get(Parent.Id);
                Google.Apis.Sheets.v4.Data.Spreadsheet objSpreadsheet = objGetSpreadsheetRequest.Execute();
                Google.Apis.Sheets.v4.Data.Sheet objSheet = objSpreadsheet.Sheets.Where(Sheet => Sheet.Properties.SheetId == SheetId).FirstOrDefault();
                SpreadsheetsResource.ValuesResource.GetRequest objGetRequest = objSheetsService.Spreadsheets.Values.Get(Parent.Id, objSheet.Properties.Title + "!" + A1NotationAddress);
                Google.Apis.Sheets.v4.Data.ValueRange objValueRange = objGetRequest.Execute();
                IList<IList<object>> GetValue = objValueRange.Values;
                object[,] arrRangeValue = new object[GetValue.Count, GetValue[0].Count];
                for (int i = 0; i < GetValue.Count; i++)
                {
                    for (int j = 0; j < GetValue[i].Count; j++)
                    {
                        arrRangeValue[i, j] = GetValue[i][j];
                    }
                }
                Values = arrRangeValue;
                return arrRangeValue;
            }
            catch
            {
                return null;
            }
        }
        public void AddFilterView(object HiddenValues = null, int FilterColumnIndex = 1, GSConditionType ConditionType = GSConditionType.gsConditionTypeUnspecified, object ConditionValues = null)
        {
            Google.Apis.Sheets.v4.SheetsService objSheetsService = (Google.Apis.Sheets.v4.SheetsService)ActiveService;
            Google.Apis.Sheets.v4.Data.Request objRequest = new Google.Apis.Sheets.v4.Data.Request();
            IList<string> colHiddenValues = new List<string>();
            if (HiddenValues != null)
            {
                Type HiddenValuesType = HiddenValues.GetType();
                if (HiddenValuesType.IsArray)
                {
                    object[] arrHiddenValues = (object[])HiddenValues;
                    for (int i = 0; i < arrHiddenValues.Length; i++)
                    {
                        colHiddenValues.Add(arrHiddenValues[i].ToString());
                    }
                }
            }
            BooleanCondition objBooleanCondition = new BooleanCondition();
            switch (ConditionType)
            {
                case GSConditionType.gsConditionTypeUnspecified:
                    objBooleanCondition.Type = "CONDITION_TYPE_UNSPECIFIED";
                    break;
                case GSConditionType.gsConditionTypeNumberGreater:
                    objBooleanCondition.Type = "NUMBER_GREATER";
                    break;
                case GSConditionType.gsConditionTypeNumberGreaterThanEQ:
                    objBooleanCondition.Type = "NUMBER_GREATER_THAN_EQ";
                    break;
                case GSConditionType.gsConditionTypeNumberLess:
                    objBooleanCondition.Type = "NUMBER_LESS";
                    break;
                case GSConditionType.gsConditionTypeNumberLessThanEQ:
                    objBooleanCondition.Type = "NUMBER_LESS_THAN_EQ";
                    break;
                case GSConditionType.gsConditionTypeNumberEQ:
                    objBooleanCondition.Type = "NUMBER_EQ";
                    break;
                case GSConditionType.gsConditionTypeNumberNotEQ:
                    objBooleanCondition.Type = "NOT_EQ";
                    break;
                case GSConditionType.gsConditionTypeNumberBetween:
                    objBooleanCondition.Type = "NUMBER_BETWEEN";
                    break;
                case GSConditionType.gsConditionTypeNumberNotBetween:
                    objBooleanCondition.Type = "NUMBER_NOT_BETWEEN";
                    break;
                case GSConditionType.gsConditionTypeTextContains:
                    objBooleanCondition.Type = "TEXT_CONTAINS";
                    break;
                case GSConditionType.gsConditionTypeTextNotContains:
                    objBooleanCondition.Type = "TEXT_NOT_CONTAINS";
                    break;
                case GSConditionType.gsConditionTypeTextStartsWith:
                    objBooleanCondition.Type = "TEXT_STARTS_WITH";
                    break;
                case GSConditionType.gsConditionTypeTextEndsWith:
                    objBooleanCondition.Type = "TEXT_ENDS_WITH";
                    break;
                case GSConditionType.gsConditionTypeTextEQ:
                    objBooleanCondition.Type = "TEXT_EQ";
                    break;
                case GSConditionType.gsConditionTypeTextIsEmail:
                    objBooleanCondition.Type = "TEXT_IS_EMAIL";
                    break;
                case GSConditionType.gsConditionTypeTextIsURL:
                    objBooleanCondition.Type = "TEXT_IS_URL";
                    break;
                case GSConditionType.gsConditionTypeDateEQ:
                    objBooleanCondition.Type = "DATE_EQ";
                    break;
                case GSConditionType.gsConditionTypeDateBefore:
                    objBooleanCondition.Type = "DATE_BEFORE";
                    break;
                case GSConditionType.gsConditionTypeDateAfter:
                    objBooleanCondition.Type = "";
                    break;
                case GSConditionType.gsConditionTypeDateOnOrBefore:
                    objBooleanCondition.Type = "DATE_ON_OR_BEFORE";
                    break;
                case GSConditionType.gsConditionTypeDateOnOrAfter:
                    objBooleanCondition.Type = "DATE_ON_OR_AFTER";
                    break;
                case GSConditionType.gsConditionTypeDateBetween:
                    objBooleanCondition.Type = "DATE_BETWEEN";
                    break;
                case GSConditionType.gsConditionTypeDateNotBetween:
                    objBooleanCondition.Type = "DATE_NOT_BETWEEN";
                    break;
                case GSConditionType.gsConditionTypeDateIsValid:
                    objBooleanCondition.Type = "DATE_IS_VALID";
                    break;
                case GSConditionType.gsConditionTypeOneOfRange:
                    objBooleanCondition.Type = "ONE_OF_RANGE";
                    break;
                case GSConditionType.gsConditionTypeOneOfList:
                    objBooleanCondition.Type = "ONE_OF_LIST";
                    break;
                case GSConditionType.gsConditionTypeBlank:
                    objBooleanCondition.Type = "BLANK";
                    break;
                case GSConditionType.gsConditionTypeNotBlank:
                    objBooleanCondition.Type = "NOT_BLANK";
                    break;
                case GSConditionType.gsConditionTypeCustomFormula:
                    objBooleanCondition.Type = "CUSTOM_FORMULA";
                    break;
                case GSConditionType.gsConditionTypeBoolean:
                    objBooleanCondition.Type = "BOOLEAN";
                    break;
                case GSConditionType.gsConditionTypeTextNotEQ:
                    objBooleanCondition.Type = "TEXT_NOT_EQ";
                    break;
                case GSConditionType.gsConditionTypeDateNotEQ:
                    objBooleanCondition.Type = "DATE_NOT_EQ";
                    break;
            }
            IList<ConditionValue> colConditionValues = new List<ConditionValue>();
            if (ConditionValues != null)
            {
                Type ConditionValuesType = ConditionValues.GetType();
                if (ConditionValuesType.IsArray)
                {
                    object[] arrConditionValues = (object[])ConditionValues;
                    for (int i = 0; i < arrConditionValues.Length; i++)
                    {
                        ConditionValue objConditionValue = new ConditionValue
                        {
                            UserEnteredValue = arrConditionValues[i].ToString(),
                        };
                        colConditionValues.Add(objConditionValue);
                    }
                    objBooleanCondition.Values = colConditionValues;
                }
            }
            FilterCriteria objFilterCriteria = new FilterCriteria
            {
                HiddenValues = colHiddenValues,
                Condition = objBooleanCondition
            };
            FilterSpec objFilterSpec = new FilterSpec
            {
                FilterCriteria = objFilterCriteria,
                ColumnIndex = FilterColumnIndex
            };
            IList<FilterSpec> colFilterSpecs = new List<FilterSpec>
            { objFilterSpec };
            objRequest.AddFilterView = new AddFilterViewRequest
            {
                Filter = new FilterView
                {
                    Range = new GridRange
                    {
                        SheetId = SheetId,
                        StartColumnIndex = StartColumnIndex - 1,
                        EndColumnIndex = EndColumnIndex,
                        StartRowIndex = StartRowIndex - 1,
                        EndRowIndex = EndRowIndex
                    },
                    FilterSpecs = colFilterSpecs
                }
            };
            Google.Apis.Sheets.v4.Data.BatchUpdateSpreadsheetRequest objBatchUpdateSpreadsheetRequest = new BatchUpdateSpreadsheetRequest
            {
                Requests = new List<Request> { objRequest }
            };
            Google.Apis.Sheets.v4.SpreadsheetsResource.BatchUpdateRequest objBatchUpdateRequest = objSheetsService.Spreadsheets.BatchUpdate(objBatchUpdateSpreadsheetRequest, Parent.Id);
            objBatchUpdateRequest.Execute();
        }
        public void Sort(GSSortOrder SortOrder, int ColumnIndex)
        {
            Google.Apis.Sheets.v4.SheetsService objSheetsService = (Google.Apis.Sheets.v4.SheetsService)ActiveService;
            Google.Apis.Sheets.v4.Data.Request objRequest = new Request();
            Google.Apis.Sheets.v4.Data.SortSpec objSortSpec = new SortSpec();
            objSortSpec.DimensionIndex = ColumnIndex;
            switch (SortOrder)
            {
                case GSSortOrder.gsSortOrderAscending:
                    objSortSpec.SortOrder = "ASCENDING";
                    break;
                case GSSortOrder.gsSortOrderDescending:
                    objSortSpec.SortOrder = "DESCENDING";
                    break;
            }
            objRequest.SortRange = new SortRangeRequest
            {
                Range = new GridRange
                {
                    SheetId = SheetId,
                    StartColumnIndex = StartColumnIndex - 1,
                    EndColumnIndex = EndColumnIndex,
                    StartRowIndex = StartRowIndex - 1,
                    EndRowIndex = EndRowIndex
                },
                SortSpecs = new List<SortSpec> { objSortSpec}
            };
            Google.Apis.Sheets.v4.Data.BatchUpdateSpreadsheetRequest objBatchUpdateSpreadsheetRequest = new BatchUpdateSpreadsheetRequest
            {
                Requests = new List<Request> { objRequest}
            };
            Google.Apis.Sheets.v4.SpreadsheetsResource.BatchUpdateRequest objBatchUpdateRequest = objSheetsService.Spreadsheets.BatchUpdate(objBatchUpdateSpreadsheetRequest, Parent.Id);
            objBatchUpdateRequest.Execute();
        }
        public void DeleteFilterView(string FilterViewName)
        {
            Google.Apis.Sheets.v4.SheetsService objSheetsService = (Google.Apis.Sheets.v4.SheetsService)ActiveService;
            Google.Apis.Sheets.v4.Data.Spreadsheet objSpreadsheet = objSheetsService.Spreadsheets.Get(Parent.Id).Execute();
            Google.Apis.Sheets.v4.Data.Sheet objSheet = objSpreadsheet.Sheets.Where(Sheet => Sheet.Properties.SheetId == SheetId).FirstOrDefault();
            Google.Apis.Sheets.v4.Data.FilterView objFilterView = objSheet.FilterViews.Where(FilterView => FilterView.Title == FilterViewName).FirstOrDefault();
            Google.Apis.Sheets.v4.Data.Request objRequest = new Request();
            objRequest.DeleteFilterView = new DeleteFilterViewRequest
            {
                FilterId = objFilterView.FilterViewId
            };
            Google.Apis.Sheets.v4.Data.BatchUpdateSpreadsheetRequest objBatchUpdateSpreadsheetRequest = new BatchUpdateSpreadsheetRequest
            {
                Requests = new List<Request> { objRequest }
            };
            Google.Apis.Sheets.v4.SpreadsheetsResource.BatchUpdateRequest objBatchUpdateRequest = objSheetsService.Spreadsheets.BatchUpdate(objBatchUpdateSpreadsheetRequest, Parent.Id);
            objBatchUpdateRequest.Execute();
        }
        public void AddConditionalFormatRule(string RuleName, object ConditionValues, object RGB, int Position = 0, GSConditionType ConditionType = GSConditionType.gsConditionTypeUnspecified)
        {
            if (Position >= 0)
            {
                Google.Apis.Sheets.v4.SheetsService objSheetsService = (Google.Apis.Sheets.v4.SheetsService)ActiveService;
                BooleanCondition objBooleanCondition = new BooleanCondition();
                switch (ConditionType)
                {
                    case GSConditionType.gsConditionTypeUnspecified:
                        objBooleanCondition.Type = "CONDITION_TYPE_UNSPECIFIED";
                        break;
                    case GSConditionType.gsConditionTypeNumberGreater:
                        objBooleanCondition.Type = "NUMBER_GREATER";
                        break;
                    case GSConditionType.gsConditionTypeNumberGreaterThanEQ:
                        objBooleanCondition.Type = "NUMBER_GREATER_THAN_EQ";
                        break;
                    case GSConditionType.gsConditionTypeNumberLess:
                        objBooleanCondition.Type = "NUMBER_LESS";
                        break;
                    case GSConditionType.gsConditionTypeNumberLessThanEQ:
                        objBooleanCondition.Type = "NUMBER_LESS_THAN_EQ";
                        break;
                    case GSConditionType.gsConditionTypeNumberEQ:
                        objBooleanCondition.Type = "NUMBER_EQ";
                        break;
                    case GSConditionType.gsConditionTypeNumberNotEQ:
                        objBooleanCondition.Type = "NOT_EQ";
                        break;
                    case GSConditionType.gsConditionTypeNumberBetween:
                        objBooleanCondition.Type = "NUMBER_BETWEEN";
                        break;
                    case GSConditionType.gsConditionTypeNumberNotBetween:
                        objBooleanCondition.Type = "NUMBER_NOT_BETWEEN";
                        break;
                    case GSConditionType.gsConditionTypeTextContains:
                        objBooleanCondition.Type = "TEXT_CONTAINS";
                        break;
                    case GSConditionType.gsConditionTypeTextNotContains:
                        objBooleanCondition.Type = "TEXT_NOT_CONTAINS";
                        break;
                    case GSConditionType.gsConditionTypeTextStartsWith:
                        objBooleanCondition.Type = "TEXT_STARTS_WITH";
                        break;
                    case GSConditionType.gsConditionTypeTextEndsWith:
                        objBooleanCondition.Type = "TEXT_ENDS_WITH";
                        break;
                    case GSConditionType.gsConditionTypeTextEQ:
                        objBooleanCondition.Type = "TEXT_EQ";
                        break;
                    case GSConditionType.gsConditionTypeTextIsEmail:
                        objBooleanCondition.Type = "TEXT_IS_EMAIL";
                        break;
                    case GSConditionType.gsConditionTypeTextIsURL:
                        objBooleanCondition.Type = "TEXT_IS_URL";
                        break;
                    case GSConditionType.gsConditionTypeDateEQ:
                        objBooleanCondition.Type = "DATE_EQ";
                        break;
                    case GSConditionType.gsConditionTypeDateBefore:
                        objBooleanCondition.Type = "DATE_BEFORE";
                        break;
                    case GSConditionType.gsConditionTypeDateAfter:
                        objBooleanCondition.Type = "";
                        break;
                    case GSConditionType.gsConditionTypeDateOnOrBefore:
                        objBooleanCondition.Type = "DATE_ON_OR_BEFORE";
                        break;
                    case GSConditionType.gsConditionTypeDateOnOrAfter:
                        objBooleanCondition.Type = "DATE_ON_OR_AFTER";
                        break;
                    case GSConditionType.gsConditionTypeDateBetween:
                        objBooleanCondition.Type = "DATE_BETWEEN";
                        break;
                    case GSConditionType.gsConditionTypeDateNotBetween:
                        objBooleanCondition.Type = "DATE_NOT_BETWEEN";
                        break;
                    case GSConditionType.gsConditionTypeDateIsValid:
                        objBooleanCondition.Type = "DATE_IS_VALID";
                        break;
                    case GSConditionType.gsConditionTypeOneOfRange:
                        objBooleanCondition.Type = "ONE_OF_RANGE";
                        break;
                    case GSConditionType.gsConditionTypeOneOfList:
                        objBooleanCondition.Type = "ONE_OF_LIST";
                        break;
                    case GSConditionType.gsConditionTypeBlank:
                        objBooleanCondition.Type = "BLANK";
                        break;
                    case GSConditionType.gsConditionTypeNotBlank:
                        objBooleanCondition.Type = "NOT_BLANK";
                        break;
                    case GSConditionType.gsConditionTypeCustomFormula:
                        objBooleanCondition.Type = "CUSTOM_FORMULA";
                        break;
                    case GSConditionType.gsConditionTypeBoolean:
                        objBooleanCondition.Type = "BOOLEAN";
                        break;
                    case GSConditionType.gsConditionTypeTextNotEQ:
                        objBooleanCondition.Type = "TEXT_NOT_EQ";
                        break;
                    case GSConditionType.gsConditionTypeDateNotEQ:
                        objBooleanCondition.Type = "DATE_NOT_EQ";
                        break;
                }
                Type RGBType = RGB.GetType();
                Color objColor = new Color();
                if (RGBType.IsArray)
                {
                    float[] arrRGB = (float[])RGB;
                    objColor.Red = arrRGB[0]/255;
                    objColor.Blue = arrRGB[1]/255;
                    objColor.Green = arrRGB[2]/255;
                    objColor.Alpha = (float)0.5;
                }
                else { throw new ArgumentException("RGB: Wrong argument type. Accepted data type: Zero-based one dimension array of variant."); }
                IList<ConditionValue> colConditionValues = new List<ConditionValue>();
                if (ConditionValues != null)
                {
                    Type ConditionValuesType = ConditionValues.GetType();
                    if (ConditionValuesType.IsArray)
                    {
                        object[] arrConditionValues = (object[])ConditionValues;
                        for (int i = 0; i < arrConditionValues.Length; i++)
                        {
                            ConditionValue objConditionValue = new ConditionValue
                            {
                                UserEnteredValue = arrConditionValues[i].ToString(),
                            };
                            colConditionValues.Add(objConditionValue);
                        }
                        objBooleanCondition.Values = colConditionValues;
                    }
                }
                GridRange objRange = new GridRange
                {
                    SheetId = SheetId,
                    StartColumnIndex = StartColumnIndex - 1,
                    EndColumnIndex = EndColumnIndex,
                    StartRowIndex = StartRowIndex - 1,
                    EndRowIndex = EndRowIndex
                };
                Google.Apis.Sheets.v4.Data.Request objRequest = new Request();
                objRequest.AddConditionalFormatRule = new AddConditionalFormatRuleRequest
                {
                    Index = Position,
                    Rule = new ConditionalFormatRule()
                    {
                        Ranges = new List<GridRange> { objRange },
                          BooleanRule = new BooleanRule
                          {
                              Condition = objBooleanCondition,
                              Format = new CellFormat
                              {
                                    BackgroundColorStyle = new ColorStyle
                                    {
                                         RgbColor = objColor
                                    }
                              }
                          }
                    }
                };
                Google.Apis.Sheets.v4.Data.BatchUpdateSpreadsheetRequest objBatchUpdateSpreadsheetRequest = new BatchUpdateSpreadsheetRequest
                {
                    Requests = new List<Request> { objRequest }
                };
                Google.Apis.Sheets.v4.SpreadsheetsResource.BatchUpdateRequest objBatchUpdateRequest = objSheetsService.Spreadsheets.BatchUpdate(objBatchUpdateSpreadsheetRequest, Parent.Id);
                objBatchUpdateRequest.Execute();
            }
            else { throw new InvalidFormatConditionIndex("The new conditional format rule position must be larger than 0 and within the conditional format rules collection"); }
        }
        public void ClearContents()
        {
            Google.Apis.Sheets.v4.SheetsService objSheetsService = (Google.Apis.Sheets.v4.SheetsService)ActiveService;
            Google.Apis.Sheets.v4.Data.ClearValuesRequest objClearValuesRequest = new ClearValuesRequest();
            Google.Apis.Sheets.v4.Data.Spreadsheet objSpreadsheet = objSheetsService.Spreadsheets.Get(Parent.Id).Execute();
            Google.Apis.Sheets.v4.Data.Sheet objSheet = objSpreadsheet.Sheets.Where(Sheet => Sheet.Properties.SheetId == SheetId).FirstOrDefault();
            SpreadsheetsResource.ValuesResource.ClearRequest objClearRequest = objSheetsService.Spreadsheets.Values.Clear(objClearValuesRequest, objSpreadsheet.SpreadsheetId, objSheet.Properties.Title + "!" +A1NotationAddress);
        }
        public void FindReplace(string Find, string Replacement, bool SearchByRegex = false, bool IncludeFormulas = false, bool MatchCase = false, bool MatchEntireCell = false)
        {
            Google.Apis.Sheets.v4.SheetsService objSheetsService = (Google.Apis.Sheets.v4.SheetsService)ActiveService;
            Google.Apis.Sheets.v4.Data.Request objRequest = new Request();
            objRequest.FindReplace = new FindReplaceRequest
            {
                Range = new GridRange
                {
                    SheetId = SheetId,
                    StartColumnIndex = StartColumnIndex - 1,
                    EndColumnIndex = EndColumnIndex,
                    StartRowIndex = StartRowIndex - 1,
                    EndRowIndex = EndRowIndex
                },
                Find = Find,
                Replacement = Replacement,
                SearchByRegex = SearchByRegex,
                IncludeFormulas = IncludeFormulas,
                MatchEntireCell = MatchEntireCell,
                MatchCase = MatchCase,
            };
            Google.Apis.Sheets.v4.Data.BatchUpdateSpreadsheetRequest objBatchUpdateSpreadsheetResource = new BatchUpdateSpreadsheetRequest
            {
                Requests = new List<Request> { objRequest}
            };
            Google.Apis.Sheets.v4.SpreadsheetsResource.BatchUpdateRequest objBatchUpdateRequest = objSheetsService.Spreadsheets.BatchUpdate(objBatchUpdateSpreadsheetResource, Parent.Id);
            objBatchUpdateRequest.Execute();
        }
    }
}
