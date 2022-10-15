using Google.Apis.Drive.v3.Data;
using Google.Apis.Sheets.v4;
using Google.Apis.Sheets.v4.Data;
using Microsoft.Vbe.Interop;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using static GoogleApis.GoogleSheetsFile;

namespace GoogleApis
{
    [Guid("C792B2AA-0ABA-4033-B492-2E8E16ECC8F4")]
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    public interface IGoogleSheetsFile
    {
        string Name { get; set; }
        string Id { get; set; }
        string[] SheetNames { get; set; }
        object ActiveService { get; set; }
        string TimeZone { get; set; }
        string AutoRecalculation { get; set; }
        string URL { get; set; }
        SheetsRange GetRange(string SheetName, string A1Range, string R1Range);
        object[,] GetRangeValues(string A1Range);
        void AddSheet(string SheetName, bool Hidden = false);
        void DeleteSheet(string SheetName);
        void CopySheet(string SheetToCopy, string NewSheetName, int Position = 0);
        void AppendValues(string A1Range, object RangeValues);
        SheetsRange AddNamedRange(string Name, string SheetName, string A1Range, string R1Range);
        void DeleteNamedRange(string Name);
        void UpdateNamedRange(string NamedRangeToUpdate, string NewName, string NewSheet, string R1Range);
        SheetsRange AddProtectedRange(string SheetName, string A1Range, string R1Range, GSEditorsType EditorsType = GSEditorsType.gsEditorsTypeNone, object Editors = null, string Description = null, bool WarningOnly = true);
        void DeleteProtectedRange(string SheetName, string R1Range);
        void UpdateProtectedRange(string SheetName, string A1Range, string R1Range, GSEditorsType EditorsType = GSEditorsType.gsEditorsTypeNone, object Editors = null, string Description = null, bool WarningOnly = true);
        void UpdateSheet(string SheetName, string NewSheetName = null, bool Hidden = false, object TabColorStyle = null, int Position = 0);
    }
    [Guid("6C84F219-EACB-44E3-BAF9-1A4221D724FE")]
    [ClassInterface(ClassInterfaceType.None)]
    public class GoogleSheetsFile : IGoogleSheetsFile
    {
        public string Name { get; set; }
        public string Id { get; set; }
        public string[] SheetNames { get; set; }
        public object ActiveService { get; set; }
        public string TimeZone { get; set; }
        public string AutoRecalculation { get; set; }
        public string URL { get; set; }
        public enum GSEditorsType
        {
            gsEditorsTypeNone,
            gsEditorsTypeUsers,
            gsEditorsTypeGroup,
            gsEditorsTypeDomain
        }
        public SheetsRange GetRange(string SheetName, string A1Range, string R1Range)
        {
            Google.Apis.Sheets.v4.SheetsService objSheetsService = (Google.Apis.Sheets.v4.SheetsService)ActiveService;
            SpreadsheetsResource.ValuesResource.GetRequest objGetRequest = objSheetsService.Spreadsheets.Values.Get(Id, SheetName + "!" + A1Range);
            Google.Apis.Sheets.v4.Data.ValueRange objValueRange = objGetRequest.Execute();
            Google.Apis.Sheets.v4.Data.Spreadsheet objSpreadsheet = objSheetsService.Spreadsheets.Get(Id).Execute();
            Google.Apis.Sheets.v4.Data.Sheet objSheet = objSpreadsheet.Sheets.Where(Sheet => Sheet.Properties.Title == SheetName).FirstOrDefault();
            int intSheetId = (int)objSheet.Properties.SheetId;
            string strR1C1NotationStyle = R1Range.Replace("R", " ");
            strR1C1NotationStyle = strR1C1NotationStyle.Replace("C", " ");
            strR1C1NotationStyle = strR1C1NotationStyle.Replace(":", "");
            string[] arrR1C1NotationAddress = strR1C1NotationStyle.Split('\u0020');
            GoogleSheetsFile objGoogleSheetsFile = OpenSheetsFile(Id);
            int intCount = strR1C1NotationStyle.Length - strR1C1NotationStyle.Replace(" ", "").Length;
            if (intCount == 4)
            {
                SheetsRange objRange = new SheetsRange
                {
                    Type = SheetsRange.GSRangeType.gsRangeTypeNormal,
                    ActiveService = ActiveService,
                    A1NotationAddress = A1Range,
                    R1C1NotationAddress = R1Range,
                    Parent = objGoogleSheetsFile,
                    SheetId = intSheetId,
                    ColumnsCount = Convert.ToInt32(arrR1C1NotationAddress[4]) - Convert.ToInt32(arrR1C1NotationAddress[2]),
                    RowsCount = Convert.ToInt32(arrR1C1NotationAddress[3]) - Convert.ToInt32(arrR1C1NotationAddress[1]),
                    StartRowIndex = Convert.ToInt32(arrR1C1NotationAddress[1]),
                    StartColumnIndex = Convert.ToInt32(arrR1C1NotationAddress[2]),
                    EndRowIndex = Convert.ToInt32(arrR1C1NotationAddress[3]),
                    EndColumnIndex = Convert.ToInt32(arrR1C1NotationAddress[4])
                };
                return objRange;
            }
            else
            {
                SheetsRange objRange = new SheetsRange
                {
                    Type = SheetsRange.GSRangeType.gsRangeTypeNormal,
                    ActiveService = ActiveService,
                    A1NotationAddress = A1Range,
                    R1C1NotationAddress = R1Range,
                    Parent = objGoogleSheetsFile,
                    SheetId = intSheetId,
                    ColumnsCount = Convert.ToInt32(arrR1C1NotationAddress[2]) - Convert.ToInt32(arrR1C1NotationAddress[2]) + 1,
                    RowsCount = Convert.ToInt32(arrR1C1NotationAddress[1]) - Convert.ToInt32(arrR1C1NotationAddress[1]) + 1,
                    StartRowIndex = Convert.ToInt32(arrR1C1NotationAddress[1]),
                    StartColumnIndex = Convert.ToInt32(arrR1C1NotationAddress[2]),
                    EndRowIndex = Convert.ToInt32(arrR1C1NotationAddress[1]),
                    EndColumnIndex = Convert.ToInt32(arrR1C1NotationAddress[2])
                };
                return objRange;
            }
        }
        public object[,] GetRangeValues(string A1Range)
        {
            try
            {
                Google.Apis.Sheets.v4.SheetsService objSheetsService = (Google.Apis.Sheets.v4.SheetsService)ActiveService;
                SpreadsheetsResource.ValuesResource.GetRequest objGetRequest = objSheetsService.Spreadsheets.Values.Get(Id, A1Range);
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
                return arrRangeValue;
            }
            catch
            {
                return null;
            }
        }
        public void AppendValues(string A1Range, object RangeValues)
        {
            Google.Apis.Sheets.v4.SheetsService objSheetsService = (Google.Apis.Sheets.v4.SheetsService)ActiveService;
            Type ValuesType = RangeValues.GetType();
            if (ValuesType.IsArray)
            {
                if (RangeValues is Array EvaluateArray && EvaluateArray.Rank == 2)
                {
                    object[,] Values2d = (object[,])RangeValues;
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
                    Google.Apis.Sheets.v4.Data.ValueRange objValueRange = new ValueRange
                    {
                        Values = ListsValues
                    };
                    SpreadsheetsResource.ValuesResource.AppendRequest objAppendRequest = objSheetsService.Spreadsheets.Values.Append(objValueRange, Id, A1Range);
                    objAppendRequest.ValueInputOption = SpreadsheetsResource.ValuesResource.AppendRequest.ValueInputOptionEnum.USERENTERED;
                    objAppendRequest.Execute();
                }
                else
                {
                    object[] Values1d = (object[])RangeValues;
                    IList<IList<object>> ListsValues = new List<IList<object>>(Values1d.GetLength(0));
                    for (int i = 0; i < Values1d.GetLength(0); i++)
                    {
                        ListsValues.Add(Values1d);
                    }
                    Google.Apis.Sheets.v4.Data.ValueRange objValueRange = new ValueRange
                    {
                        Values = ListsValues
                    };
                    SpreadsheetsResource.ValuesResource.AppendRequest objAppendRequest = objSheetsService.Spreadsheets.Values.Append(objValueRange, Id, A1Range);
                    objAppendRequest.ValueInputOption = SpreadsheetsResource.ValuesResource.AppendRequest.ValueInputOptionEnum.USERENTERED;
                    objAppendRequest.Execute();
                }
            }
            else
            {
                throw new ArgumentException("Invalid argument data type. Accepted data type: Zero-based two dimension array of variant.");
            }
        }
        private GoogleSheetsFile OpenSheetsFile(string SheetsFileId)
        {
            Google.Apis.Sheets.v4.SheetsService objSheetsService = (Google.Apis.Sheets.v4.SheetsService)ActiveService;
            Google.Apis.Sheets.v4.SpreadsheetsResource.GetRequest objGetRequest = objSheetsService.Spreadsheets.Get(SheetsFileId);
            Google.Apis.Sheets.v4.Data.Spreadsheet objSpreadsheet = objGetRequest.Execute();
            return InstantiateNewGoogleSheetsFileObject(objSpreadsheet);
        }
        private GoogleSheetsFile InstantiateNewGoogleSheetsFileObject(Google.Apis.Sheets.v4.Data.Spreadsheet Spreadsheet)
        {
            GoogleSheetsFile objSheetsFile = new GoogleSheetsFile
            {
                ActiveService = ActiveService,
                Id = Spreadsheet.SpreadsheetId,
                Name = Spreadsheet.Properties.Title,
                TimeZone = Spreadsheet.Properties.TimeZone,
                AutoRecalculation = Spreadsheet.Properties.AutoRecalc,
                SheetNames = GetSheetNamesInSpreadsheet(Spreadsheet),
                URL = Spreadsheet.SpreadsheetUrl
            };
            return objSheetsFile;
        }
        private string[] GetSheetNamesInSpreadsheet(Google.Apis.Sheets.v4.Data.Spreadsheet Spreadsheet)
        {
            string[] arrSheetNames = new string[Spreadsheet.Sheets.Count];
            int i = 0;
            foreach (Google.Apis.Sheets.v4.Data.Sheet objSheet in Spreadsheet.Sheets)
            {
                arrSheetNames[i] = objSheet.Properties.Title;
                i++;
            }
            return arrSheetNames;
        }
        public void AddSheet(string SheetName, bool Hidden = false)
        {
            Google.Apis.Sheets.v4.SheetsService objSheetsService = (Google.Apis.Sheets.v4.SheetsService)ActiveService;
            Google.Apis.Sheets.v4.Data.Request objRequest = new Request();
            objRequest.AddSheet = new AddSheetRequest
            {
                Properties = new Google.Apis.Sheets.v4.Data.SheetProperties
                {
                    Title = SheetName,
                    Hidden = Hidden
                }
            };
            Google.Apis.Sheets.v4.Data.BatchUpdateSpreadsheetRequest objBatchUpdateSpreadsheetRequest = new BatchUpdateSpreadsheetRequest
            {
                Requests = new List<Request> { objRequest }
            };
            Google.Apis.Sheets.v4.SpreadsheetsResource.BatchUpdateRequest objBatchUpdateRequest = objSheetsService.Spreadsheets.BatchUpdate(objBatchUpdateSpreadsheetRequest, Id);
            objBatchUpdateRequest.Execute();
        }
        public void DeleteSheet(string SheetName)
        {
            Google.Apis.Sheets.v4.SheetsService objSheetsService = (Google.Apis.Sheets.v4.SheetsService)ActiveService;
            Google.Apis.Sheets.v4.Data.Request objRequest = new Request();
            Google.Apis.Sheets.v4.SpreadsheetsResource.GetRequest objGetRequest = objSheetsService.Spreadsheets.Get(Id);
            Google.Apis.Sheets.v4.Data.Spreadsheet objSpreadsheet = objGetRequest.Execute();
            Google.Apis.Sheets.v4.Data.Sheet objSheet = objSpreadsheet.Sheets.Where(Sheet => Sheet.Properties.Title == SheetName).FirstOrDefault();
            objRequest.DeleteSheet = new DeleteSheetRequest
            {
                SheetId = objSheet.Properties.SheetId
            };
            Google.Apis.Sheets.v4.Data.BatchUpdateSpreadsheetRequest objBatchUpdateSpreadsheetRequest = new BatchUpdateSpreadsheetRequest
            {
                Requests = new List<Request> { objRequest}
            };
            Google.Apis.Sheets.v4.SpreadsheetsResource.BatchUpdateRequest objBatchUpdateRequest = objSheetsService.Spreadsheets.BatchUpdate(objBatchUpdateSpreadsheetRequest, Id);
            objBatchUpdateRequest.Execute();
        }
        public void CopySheet(string SheetToCopy, string NewSheetName, int Position = 0)
        {
            Google.Apis.Sheets.v4.SheetsService objSheetsService = (Google.Apis.Sheets.v4.SheetsService)ActiveService;
            Google.Apis.Sheets.v4.Data.Spreadsheet objSpreadsheet = objSheetsService.Spreadsheets.Get(Id).Execute();
            if (Position >= 0 && Position < objSpreadsheet.Sheets.Count)
            {
                Google.Apis.Sheets.v4.Data.Request objRequest = new Request();
                Google.Apis.Sheets.v4.Data.Sheet objSheet = objSpreadsheet.Sheets.Where(Sheet => Sheet.Properties.Title == SheetToCopy).FirstOrDefault();
                objRequest.DuplicateSheet = new DuplicateSheetRequest
                {
                    NewSheetName = NewSheetName,
                    SourceSheetId = objSheet.Properties.SheetId,
                    InsertSheetIndex = Position
                };
                Google.Apis.Sheets.v4.Data.BatchUpdateSpreadsheetRequest objBatchUpdateSpreadsheetRequest = new BatchUpdateSpreadsheetRequest
                {
                    Requests = new List<Request> { objRequest }
                };
                Google.Apis.Sheets.v4.SpreadsheetsResource.BatchUpdateRequest objBatchUpdateRequest = objSheetsService.Spreadsheets.BatchUpdate(objBatchUpdateSpreadsheetRequest, Id);
                objBatchUpdateRequest.Execute();
            }
            else { throw new InvalidSheetIndex("The new sheet position must be larger than 0 and within the sheets collection"); }
        }
        public SheetsRange AddNamedRange(string Name, string SheetName, string A1Range, string R1Range)
        {
            Google.Apis.Sheets.v4.SheetsService objSheetsService = (Google.Apis.Sheets.v4.SheetsService)ActiveService;
            SpreadsheetsResource.ValuesResource.GetRequest objGetRequest = objSheetsService.Spreadsheets.Values.Get(Id, SheetName + "!" + A1Range);
            Google.Apis.Sheets.v4.Data.ValueRange objValueRange = objGetRequest.Execute();
            Google.Apis.Sheets.v4.Data.Spreadsheet objSpreadsheet = objSheetsService.Spreadsheets.Get(Id).Execute();
            Google.Apis.Sheets.v4.Data.Sheet objSheet = objSpreadsheet.Sheets.Where(Sheet => Sheet.Properties.Title == SheetName).FirstOrDefault();
            int intSheetId = (int)objSheet.Properties.SheetId;
            string strR1C1NotationStyle = R1Range.Replace("R", " ");
            strR1C1NotationStyle = strR1C1NotationStyle.Replace("C", " ");
            strR1C1NotationStyle = strR1C1NotationStyle.Replace(":", "");
            string[] arrR1C1NotationAddress = strR1C1NotationStyle.Split('\u0020');
            Google.Apis.Sheets.v4.Data.Request objRequest = new Request();
            Google.Apis.Sheets.v4.Data.GridRange objGridRange = new GridRange();
            int intCount = strR1C1NotationStyle.Length - strR1C1NotationStyle.Replace(" ", "").Length;
            if (intCount != 4)
            {
                objGridRange.SheetId = intSheetId;
                objGridRange.StartColumnIndex = Convert.ToInt32(arrR1C1NotationAddress[2]) - 1;
                objGridRange.EndColumnIndex = Convert.ToInt32(arrR1C1NotationAddress[2]);
                objGridRange.StartRowIndex = Convert.ToInt32(arrR1C1NotationAddress[1]) - 1;
                objGridRange.EndRowIndex = Convert.ToInt32(arrR1C1NotationAddress[1]);
            }
            else
            {
                objGridRange.SheetId = intSheetId;
                objGridRange.StartColumnIndex = Convert.ToInt32(arrR1C1NotationAddress[2]) - 1;
                objGridRange.EndColumnIndex = Convert.ToInt32(arrR1C1NotationAddress[4]);
                objGridRange.StartRowIndex = Convert.ToInt32(arrR1C1NotationAddress[1]) - 1;
                objGridRange.EndRowIndex = Convert.ToInt32(arrR1C1NotationAddress[3]);
            }
            objRequest.AddNamedRange = new AddNamedRangeRequest
            {
                NamedRange = new NamedRange
                {
                    Name = Name,
                    Range = objGridRange
                }
            };
            Google.Apis.Sheets.v4.Data.BatchUpdateSpreadsheetRequest objBatchUpdateSpreadsheetRequest = new BatchUpdateSpreadsheetRequest
            {
                Requests = new List<Request> { objRequest }
            };
            Google.Apis.Sheets.v4.SpreadsheetsResource.BatchUpdateRequest objBatchUpdateRequest = objSheetsService.Spreadsheets.BatchUpdate(objBatchUpdateSpreadsheetRequest, Id);
            objBatchUpdateRequest.Execute();
            GoogleSheetsFile objGoogleSheetsFile = OpenSheetsFile(Id);
            if (intCount == 4)
            {
                SheetsRange objRange = new SheetsRange
                {
                    Type = SheetsRange.GSRangeType.gsRangeTypeNormal,
                    ActiveService = ActiveService,
                    A1NotationAddress = A1Range,
                    R1C1NotationAddress = R1Range,
                    Parent = objGoogleSheetsFile,
                    SheetId = intSheetId,
                    ColumnsCount = Convert.ToInt32(arrR1C1NotationAddress[4]) - Convert.ToInt32(arrR1C1NotationAddress[2]),
                    RowsCount = Convert.ToInt32(arrR1C1NotationAddress[3]) - Convert.ToInt32(arrR1C1NotationAddress[1]),
                    StartRowIndex = Convert.ToInt32(arrR1C1NotationAddress[1]),
                    StartColumnIndex = Convert.ToInt32(arrR1C1NotationAddress[2]),
                    EndRowIndex = Convert.ToInt32(arrR1C1NotationAddress[3]),
                    EndColumnIndex = Convert.ToInt32(arrR1C1NotationAddress[4])
                };
                return objRange;
            }
            else
            {
                SheetsRange objRange = new SheetsRange
                {
                    Type = SheetsRange.GSRangeType.gsRangeTypeNormal,
                    ActiveService = ActiveService,
                    A1NotationAddress = A1Range,
                    R1C1NotationAddress = R1Range,
                    Parent = objGoogleSheetsFile,
                    SheetId = intSheetId,
                    ColumnsCount = Convert.ToInt32(arrR1C1NotationAddress[2]) - Convert.ToInt32(arrR1C1NotationAddress[2]) + 1,
                    RowsCount = Convert.ToInt32(arrR1C1NotationAddress[1]) - Convert.ToInt32(arrR1C1NotationAddress[1]) + 1,
                    StartRowIndex = Convert.ToInt32(arrR1C1NotationAddress[1]),
                    StartColumnIndex = Convert.ToInt32(arrR1C1NotationAddress[2]),
                    EndRowIndex = Convert.ToInt32(arrR1C1NotationAddress[1]),
                    EndColumnIndex = Convert.ToInt32(arrR1C1NotationAddress[2])
                };
                return objRange;
            }
        }
        public void DeleteNamedRange(string Name)
        {
            Google.Apis.Sheets.v4.SheetsService objSheetsService = (Google.Apis.Sheets.v4.SheetsService)ActiveService;
            Google.Apis.Sheets.v4.Data.Spreadsheet objSpreadsheet = objSheetsService.Spreadsheets.Get(Id).Execute();
            Google.Apis.Sheets.v4.Data.NamedRange objNamedRange = objSpreadsheet.NamedRanges.Where(NamedRange => NamedRange.Name == Name).FirstOrDefault();
            Google.Apis.Sheets.v4.Data.Request objRequest = new Request();
            objRequest.DeleteNamedRange = new DeleteNamedRangeRequest
            {
                NamedRangeId = objNamedRange.NamedRangeId
            };
            Google.Apis.Sheets.v4.Data.BatchUpdateSpreadsheetRequest objBatchUpdateSpreadsheetRequest = new BatchUpdateSpreadsheetRequest
            {
                Requests = new List<Request> { objRequest }
            };
            Google.Apis.Sheets.v4.SpreadsheetsResource.BatchUpdateRequest objBatchUpdateRequest = objSheetsService.Spreadsheets.BatchUpdate(objBatchUpdateSpreadsheetRequest, Id);
            objBatchUpdateRequest.Execute();
        }
        public void UpdateNamedRange(string NamedRangeToUpdate, string NewName, string NewSheet, string R1Range)
        {
            Google.Apis.Sheets.v4.SheetsService objSheetsService = (Google.Apis.Sheets.v4.SheetsService)ActiveService;
            Google.Apis.Sheets.v4.Data.Spreadsheet objSpreadsheet = objSheetsService.Spreadsheets.Get(Id).Execute();
            Google.Apis.Sheets.v4.Data.NamedRange objNamedRange = objSpreadsheet.NamedRanges.Where(NamedRange => NamedRange.Name == NamedRangeToUpdate).FirstOrDefault();
            Google.Apis.Sheets.v4.Data.Sheet objSheet = objSpreadsheet.Sheets.Where(Sheet => Sheet.Properties.Title == NewSheet).FirstOrDefault();
            string strR1C1NotationStyle = R1Range.Replace("R", " ");
            strR1C1NotationStyle = strR1C1NotationStyle.Replace("C", " ");
            strR1C1NotationStyle = strR1C1NotationStyle.Replace(":", "");
            string[] arrR1C1NotationAddress = strR1C1NotationStyle.Split('\u0020');
            Google.Apis.Sheets.v4.Data.Request objRequest = new Request();
            Google.Apis.Sheets.v4.Data.GridRange objGridRange = new GridRange();
            int intCount = strR1C1NotationStyle.Length - strR1C1NotationStyle.Replace(" ", "").Length;
            if (intCount != 4)
            {
                objGridRange.SheetId = objSheet.Properties.SheetId;
                objGridRange.StartColumnIndex = Convert.ToInt32(arrR1C1NotationAddress[2]) - 1;
                objGridRange.EndColumnIndex = Convert.ToInt32(arrR1C1NotationAddress[2]);
                objGridRange.StartRowIndex = Convert.ToInt32(arrR1C1NotationAddress[1]) - 1;
                objGridRange.EndRowIndex = Convert.ToInt32(arrR1C1NotationAddress[1]);
            }
            else
            {
                objGridRange.SheetId = objSheet.Properties.SheetId;
                objGridRange.StartColumnIndex = Convert.ToInt32(arrR1C1NotationAddress[2]) - 1;
                objGridRange.EndColumnIndex = Convert.ToInt32(arrR1C1NotationAddress[4]);
                objGridRange.StartRowIndex = Convert.ToInt32(arrR1C1NotationAddress[1]) - 1;
                objGridRange.EndRowIndex = Convert.ToInt32(arrR1C1NotationAddress[3]);
            }
            objRequest.UpdateNamedRange = new UpdateNamedRangeRequest
            {
                NamedRange = new NamedRange
                {
                    Name = NewName,
                    NamedRangeId = objNamedRange.NamedRangeId,
                    Range = objGridRange
                },
                Fields = "*"
            };
            Google.Apis.Sheets.v4.Data.BatchUpdateSpreadsheetRequest objBatchUpdateSpreadsheetRequest = new BatchUpdateSpreadsheetRequest
            {
                Requests = new List<Request> { objRequest }
            };
            Google.Apis.Sheets.v4.SpreadsheetsResource.BatchUpdateRequest objBatchUpdateRequest = objSheetsService.Spreadsheets.BatchUpdate(objBatchUpdateSpreadsheetRequest, Id);
            objBatchUpdateRequest.Execute();
        }
        public SheetsRange AddProtectedRange(string SheetName, string A1Range, string R1Range, GSEditorsType EditorsType = GSEditorsType.gsEditorsTypeNone, object Editors = null, string Description = null, bool WarningOnly = true)
        {
            Google.Apis.Sheets.v4.SheetsService objSheetsService = (Google.Apis.Sheets.v4.SheetsService)ActiveService;
            SpreadsheetsResource.ValuesResource.GetRequest objGetRequest = objSheetsService.Spreadsheets.Values.Get(Id, SheetName + "!" + A1Range);
            Google.Apis.Sheets.v4.Data.ValueRange objValueRange = objGetRequest.Execute();
            Google.Apis.Sheets.v4.Data.Spreadsheet objSpreadsheet = objSheetsService.Spreadsheets.Get(Id).Execute();
            Google.Apis.Sheets.v4.Data.Sheet objSheet = objSpreadsheet.Sheets.Where(Sheet => Sheet.Properties.Title == SheetName).FirstOrDefault();
            int intSheetId = (int)objSheet.Properties.SheetId;
            string strR1C1NotationStyle = R1Range.Replace("R", " ");
            strR1C1NotationStyle = strR1C1NotationStyle.Replace("C", " ");
            strR1C1NotationStyle = strR1C1NotationStyle.Replace(":", "");
            string[] arrR1C1NotationAddress = strR1C1NotationStyle.Split('\u0020');
            Google.Apis.Sheets.v4.Data.GridRange objGridRange = new GridRange();
            int intCount = strR1C1NotationStyle.Length - strR1C1NotationStyle.Replace(" ", "").Length;
            if (intCount != 4)
            {
                objGridRange.SheetId = intSheetId;
                objGridRange.StartColumnIndex = Convert.ToInt32(arrR1C1NotationAddress[2]) - 1;
                objGridRange.EndColumnIndex = Convert.ToInt32(arrR1C1NotationAddress[2]);
                objGridRange.StartRowIndex = Convert.ToInt32(arrR1C1NotationAddress[1]) - 1;
                objGridRange.EndRowIndex = Convert.ToInt32(arrR1C1NotationAddress[1]);
            }
            else
            {
                objGridRange.SheetId = intSheetId;
                objGridRange.StartColumnIndex = Convert.ToInt32(arrR1C1NotationAddress[2]) - 1;
                objGridRange.EndColumnIndex = Convert.ToInt32(arrR1C1NotationAddress[4]);
                objGridRange.StartRowIndex = Convert.ToInt32(arrR1C1NotationAddress[1]) - 1;
                objGridRange.EndRowIndex = Convert.ToInt32(arrR1C1NotationAddress[3]);
            }
            Google.Apis.Sheets.v4.Data.ProtectedRange objProtectedRange = new ProtectedRange
            {
                Description = Description,
                WarningOnly = WarningOnly,
                Range = objGridRange
            };
            List<string> listEditors = new List<string>();
            Google.Apis.Sheets.v4.Data.Editors objEditors = new Editors();
            if (WarningOnly == false && EditorsType != GSEditorsType.gsEditorsTypeNone)
            {
                if (Editors != null)
                {
                    Type GetEditorsType = Editors.GetType();
                    if (GetEditorsType.IsArray)
                    {
                        string[] arrEditors = (string[])Editors;
                        listEditors = arrEditors.ToList();
                        switch (EditorsType)
                        {
                            case GSEditorsType.gsEditorsTypeUsers:
                                objEditors.Users = listEditors;
                                objProtectedRange.Editors = objEditors;
                                break;
                            case GSEditorsType.gsEditorsTypeGroup:
                                objEditors.Groups = listEditors;
                                objProtectedRange.Editors = objEditors;
                                break;
                            case GSEditorsType.gsEditorsTypeDomain:
                                objEditors.DomainUsersCanEdit = true;
                                objProtectedRange.Editors = objEditors;
                                break;
                        }
                    }
                }
            }
            else { throw new InvalidEditorsType("Cannot add users or groups to the protected range when WarningOnly = true"); }
            Google.Apis.Sheets.v4.Data.Request objRequest = new Request();
            objRequest.AddProtectedRange = new AddProtectedRangeRequest
            {
                ProtectedRange = objProtectedRange
            };
            Google.Apis.Sheets.v4.Data.BatchUpdateSpreadsheetRequest objBatchUpdateSpreadsheetRequest = new BatchUpdateSpreadsheetRequest
            {
                Requests = new List<Request> { objRequest }
            };
            Google.Apis.Sheets.v4.SpreadsheetsResource.BatchUpdateRequest objBatchUpdateRequest = objSheetsService.Spreadsheets.BatchUpdate(objBatchUpdateSpreadsheetRequest, Id);
            objBatchUpdateRequest.Execute();
            GoogleSheetsFile objGoogleSheetsFile = OpenSheetsFile(Id);
            if (intCount == 4)
            {
                SheetsRange objRange = new SheetsRange
                {
                    Type = SheetsRange.GSRangeType.gsRangeTypeNormal,
                    ActiveService = ActiveService,
                    A1NotationAddress = A1Range,
                    R1C1NotationAddress = R1Range,
                    Parent = objGoogleSheetsFile,
                    SheetId = intSheetId,
                    ColumnsCount = Convert.ToInt32(arrR1C1NotationAddress[4]) - Convert.ToInt32(arrR1C1NotationAddress[2]),
                    RowsCount = Convert.ToInt32(arrR1C1NotationAddress[3]) - Convert.ToInt32(arrR1C1NotationAddress[1]),
                    StartRowIndex = Convert.ToInt32(arrR1C1NotationAddress[1]),
                    StartColumnIndex = Convert.ToInt32(arrR1C1NotationAddress[2]),
                    EndRowIndex = Convert.ToInt32(arrR1C1NotationAddress[3]),
                    EndColumnIndex = Convert.ToInt32(arrR1C1NotationAddress[4])
                };
                return objRange;
            }
            else
            {
                SheetsRange objRange = new SheetsRange
                {
                    Type = SheetsRange.GSRangeType.gsRangeTypeNormal,
                    ActiveService = ActiveService,
                    A1NotationAddress = A1Range,
                    R1C1NotationAddress = R1Range,
                    Parent = objGoogleSheetsFile,
                    SheetId = intSheetId,
                    ColumnsCount = Convert.ToInt32(arrR1C1NotationAddress[2]) - Convert.ToInt32(arrR1C1NotationAddress[2]) + 1,
                    RowsCount = Convert.ToInt32(arrR1C1NotationAddress[1]) - Convert.ToInt32(arrR1C1NotationAddress[1]) + 1,
                    StartRowIndex = Convert.ToInt32(arrR1C1NotationAddress[1]),
                    StartColumnIndex = Convert.ToInt32(arrR1C1NotationAddress[2]),
                    EndRowIndex = Convert.ToInt32(arrR1C1NotationAddress[1]),
                    EndColumnIndex = Convert.ToInt32(arrR1C1NotationAddress[2])
                };
                return objRange;
            }
        }
        public void UpdateProtectedRange(string SheetName, string A1Range, string R1Range, GSEditorsType EditorsType = GSEditorsType.gsEditorsTypeNone, object Editors = null, string Description = null, bool WarningOnly = true)
        {
            Google.Apis.Sheets.v4.SheetsService objSheetsService = (Google.Apis.Sheets.v4.SheetsService)ActiveService;
            SpreadsheetsResource.ValuesResource.GetRequest objGetRequest = objSheetsService.Spreadsheets.Values.Get(Id, SheetName + "!" + A1Range);
            Google.Apis.Sheets.v4.Data.ValueRange objValueRange = objGetRequest.Execute();
            Google.Apis.Sheets.v4.Data.Spreadsheet objSpreadsheet = objSheetsService.Spreadsheets.Get(Id).Execute();
            Google.Apis.Sheets.v4.Data.Sheet objSheet = objSpreadsheet.Sheets.Where(Sheet => Sheet.Properties.Title == SheetName).FirstOrDefault();
            int intSheetId = (int)objSheet.Properties.SheetId;
            string strR1C1NotationStyle = R1Range.Replace("R", " ");
            strR1C1NotationStyle = strR1C1NotationStyle.Replace("C", " ");
            strR1C1NotationStyle = strR1C1NotationStyle.Replace(":", "");
            string[] arrR1C1NotationAddress = strR1C1NotationStyle.Split('\u0020');
            Google.Apis.Sheets.v4.Data.GridRange objGridRange = new GridRange();
            int intCount = strR1C1NotationStyle.Length - strR1C1NotationStyle.Replace(" ", "").Length;
            if (intCount != 4)
            {
                objGridRange.SheetId = intSheetId;
                objGridRange.StartColumnIndex = Convert.ToInt32(arrR1C1NotationAddress[2]) - 1;
                objGridRange.EndColumnIndex = Convert.ToInt32(arrR1C1NotationAddress[2]);
                objGridRange.StartRowIndex = Convert.ToInt32(arrR1C1NotationAddress[1]) - 1;
                objGridRange.EndRowIndex = Convert.ToInt32(arrR1C1NotationAddress[1]);
            }
            else
            {
                objGridRange.SheetId = intSheetId;
                objGridRange.StartColumnIndex = Convert.ToInt32(arrR1C1NotationAddress[2]) - 1;
                objGridRange.EndColumnIndex = Convert.ToInt32(arrR1C1NotationAddress[4]);
                objGridRange.StartRowIndex = Convert.ToInt32(arrR1C1NotationAddress[1]) - 1;
                objGridRange.EndRowIndex = Convert.ToInt32(arrR1C1NotationAddress[3]);
            }
            Google.Apis.Sheets.v4.Data.ProtectedRange objProtectedRange = new ProtectedRange
            {
                Description = Description,
                WarningOnly = WarningOnly,
                Range = objGridRange
            };
            List<string> listEditors = new List<string>();
            Google.Apis.Sheets.v4.Data.Editors objEditors = new Editors();
            if (WarningOnly == false && EditorsType != GSEditorsType.gsEditorsTypeNone)
            {
                if (Editors != null)
                {
                    Type GetEditorsType = Editors.GetType();
                    if (GetEditorsType.IsArray)
                    {
                        string[] arrEditors = (string[])Editors;
                        listEditors = arrEditors.ToList();
                        switch (EditorsType)
                        {
                            case GSEditorsType.gsEditorsTypeUsers:
                                objEditors.Users = listEditors;
                                objProtectedRange.Editors = objEditors;
                                break;
                            case GSEditorsType.gsEditorsTypeGroup:
                                objEditors.Groups = listEditors;
                                objProtectedRange.Editors = objEditors;
                                break;
                            case GSEditorsType.gsEditorsTypeDomain:
                                objEditors.DomainUsersCanEdit = true;
                                objProtectedRange.Editors = objEditors;
                                break;
                        }
                    }
                }
            }
            else { throw new InvalidEditorsType("Cannot add users or groups to the protected range when WarningOnly = true"); }
            Google.Apis.Sheets.v4.Data.Request objRequest = new Request();
            objRequest.UpdateProtectedRange = new UpdateProtectedRangeRequest
            {
                ProtectedRange = objProtectedRange,
                Fields = "*"
            };
            Google.Apis.Sheets.v4.Data.BatchUpdateSpreadsheetRequest objBatchUpdateSpreadsheetRequest = new BatchUpdateSpreadsheetRequest
            {
                Requests = new List<Request> { objRequest }
            };
            Google.Apis.Sheets.v4.SpreadsheetsResource.BatchUpdateRequest objBatchUpdateRequest = objSheetsService.Spreadsheets.BatchUpdate(objBatchUpdateSpreadsheetRequest, Id);
            objBatchUpdateRequest.Execute();
        }
        public void DeleteProtectedRange(string SheetName, string R1Range)
        {
            Google.Apis.Sheets.v4.SheetsService objSheetsService = (Google.Apis.Sheets.v4.SheetsService)ActiveService;
            Google.Apis.Sheets.v4.Data.Spreadsheet objSpreadsheet = objSheetsService.Spreadsheets.Get(Id).Execute();
            Google.Apis.Sheets.v4.Data.Sheet objSheet = objSpreadsheet.Sheets.Where(Sheet => Sheet.Properties.Title == SheetName).FirstOrDefault();
            string strR1C1NotationStyle = R1Range.Replace("R", " ");
            strR1C1NotationStyle = strR1C1NotationStyle.Replace("C", " ");
            strR1C1NotationStyle = strR1C1NotationStyle.Replace(":", "");
            string[] arrR1C1NotationAddress = strR1C1NotationStyle.Split('\u0020');
            Google.Apis.Sheets.v4.Data.GridRange objGridRange = new GridRange();
            int intCount = strR1C1NotationStyle.Length - strR1C1NotationStyle.Replace(" ", "").Length;
            if (intCount != 4)
            {
                objGridRange.SheetId = objSheet.Properties.SheetId;
                objGridRange.StartColumnIndex = Convert.ToInt32(arrR1C1NotationAddress[2]) - 1;
                objGridRange.EndColumnIndex = Convert.ToInt32(arrR1C1NotationAddress[2]);
                objGridRange.StartRowIndex = Convert.ToInt32(arrR1C1NotationAddress[1]) - 1;
                objGridRange.EndRowIndex = Convert.ToInt32(arrR1C1NotationAddress[1]);
            }
            else
            {
                objGridRange.SheetId = objSheet.Properties.SheetId;
                objGridRange.StartColumnIndex = Convert.ToInt32(arrR1C1NotationAddress[2]) - 1;
                objGridRange.EndColumnIndex = Convert.ToInt32(arrR1C1NotationAddress[4]);
                objGridRange.StartRowIndex = Convert.ToInt32(arrR1C1NotationAddress[1]) - 1;
                objGridRange.EndRowIndex = Convert.ToInt32(arrR1C1NotationAddress[3]);
            }
            Google.Apis.Sheets.v4.Data.ProtectedRange objProtectedRange = objSheet.ProtectedRanges
                .Where(ProtectedRange => ProtectedRange.Range.StartColumnIndex == objGridRange.StartColumnIndex 
                && ProtectedRange.Range.EndColumnIndex == objGridRange.EndColumnIndex
                && ProtectedRange.Range.StartRowIndex == objGridRange.StartRowIndex
                && ProtectedRange.Range.EndRowIndex == objGridRange.EndRowIndex
                && ProtectedRange.Range.SheetId == objGridRange.SheetId).FirstOrDefault();
            Google.Apis.Sheets.v4.Data.Request objRequest = new Request();
            objRequest.DeleteProtectedRange = new DeleteProtectedRangeRequest
            {
                ProtectedRangeId = objProtectedRange.ProtectedRangeId
            };
            Google.Apis.Sheets.v4.Data.BatchUpdateSpreadsheetRequest objBatchUpdateSpreadsheetRequest = new BatchUpdateSpreadsheetRequest
            {
                Requests = new List<Request> { objRequest }
            };
            Google.Apis.Sheets.v4.SpreadsheetsResource.BatchUpdateRequest objBatchUpdateRequest = objSheetsService.Spreadsheets.BatchUpdate(objBatchUpdateSpreadsheetRequest, Id);
            objBatchUpdateRequest.Execute();
        }
        public void UpdateSheet(string SheetName, string NewSheetName = null, bool Hidden = false, object TabColorStyle = null, int Position = 0)
        {
            Google.Apis.Sheets.v4.SheetsService objSheetsService = (Google.Apis.Sheets.v4.SheetsService)ActiveService;
            Google.Apis.Sheets.v4.Data.Spreadsheet objSpreadsheet = objSheetsService.Spreadsheets.Get(Id).Execute();
            Google.Apis.Sheets.v4.Data.Sheet objSheet = objSpreadsheet.Sheets.Where(Sheet => Sheet.Properties.Title == SheetName).FirstOrDefault();
            Google.Apis.Sheets.v4.Data.Request objRequest = new Request();
            Google.Apis.Sheets.v4.Data.SheetProperties objSheetProperties = new SheetProperties();
            Type RGBType = TabColorStyle.GetType();
            Color objColor = new Color();
            if (RGBType.IsArray)
            {
                float[] arrRGB = (float[])TabColorStyle;
                objColor.Red = arrRGB[0] / 255;
                objColor.Blue = arrRGB[1] / 255;
                objColor.Green = arrRGB[2] / 255;
                objColor.Alpha = (float)0.5;
                objSheetProperties.TabColorStyle = new ColorStyle
                {
                    RgbColor = objColor
                };
            }
            if (Position >= 0 && Position < objSpreadsheet.Sheets.Count) { objSheetProperties.Index = Position; }
            else { throw new InvalidSheetIndex("The new sheet position must be larger than 0 and within the sheets collection"); }
            if (NewSheetName != null) { objSheetProperties.Title = NewSheetName; }
            objSheetProperties.Hidden = Hidden;
            objSheetProperties.SheetId = objSheet.Properties.SheetId;
            objRequest.UpdateSheetProperties = new UpdateSheetPropertiesRequest
            {
                Fields = "title, hidden, tabColorStyle, index",
                Properties = objSheetProperties
            };
            Google.Apis.Sheets.v4.Data.BatchUpdateSpreadsheetRequest objBatchUpdateSpreadsheetRequest = new BatchUpdateSpreadsheetRequest
            {
                Requests = new List<Request> { objRequest }
            };
            Google.Apis.Sheets.v4.SpreadsheetsResource.BatchUpdateRequest objBatchUpdateRequest = objSheetsService.Spreadsheets.BatchUpdate(objBatchUpdateSpreadsheetRequest, Id);
            objBatchUpdateRequest.Execute();
        }
    }
}
