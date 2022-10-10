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
        SheetsRange GetRange(string SheetName, string A1NotationStyle, string R1C1NotationStyle);
        object[,] GetRangeValues(string A1Range);
        void AddSheet(string SheetName, bool Hidden = false);
        void DeleteSheet(string SheetName);
        void CopySheet(string SheetToCopy, string NewSheetName, int Position = 0);
        void AppendValues(string A1Range, object RangeValues);
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
        public SheetsRange GetRange(string SheetName, string A1NotationStyle, string R1C1NotationStyle)
        {
            Google.Apis.Sheets.v4.SheetsService objSheetsService = (Google.Apis.Sheets.v4.SheetsService)ActiveService;
            SpreadsheetsResource.ValuesResource.GetRequest objGetRequest = objSheetsService.Spreadsheets.Values.Get(Id, SheetName + "!" + A1NotationStyle);
            Google.Apis.Sheets.v4.Data.ValueRange objValueRange = objGetRequest.Execute();
            Google.Apis.Sheets.v4.Data.Spreadsheet objSpreadsheet = objSheetsService.Spreadsheets.Get(Id).Execute();
            Google.Apis.Sheets.v4.Data.Sheet objSheet = objSpreadsheet.Sheets.Where(Sheet => Sheet.Properties.Title == SheetName).FirstOrDefault();
            int intSheetId = (int)objSheet.Properties.SheetId;
            string strR1C1NotationStyle = R1C1NotationStyle.Replace("R", " ");
            strR1C1NotationStyle = strR1C1NotationStyle.Replace("C", " ");
            strR1C1NotationStyle = strR1C1NotationStyle.Replace(":", "");
            string[] arrR1C1NotationAddress = strR1C1NotationStyle.Split('\u0020');
            GoogleSheetsFile objGoogleSheetsFile = OpenSheetsFile(Id);
            SheetsRange objRange = new SheetsRange
            {
                Values = ConvertListOfListsTo2dArray(objValueRange.Values),
                ActiveService = ActiveService,
                A1NotationAddress = A1NotationStyle,
                R1C1NotationAddress = R1C1NotationStyle,
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
        private object[,] ConvertListOfListsTo2dArray(IList<IList<object>>ListsObject)
        {
            if (ListsObject != null) 
            { 
                object[,] arrRangeValue = new object[ListsObject.Count, ListsObject[0].Count];
                for (int i = 0; i < ListsObject.Count; i++)
                {
                    for (int j = 0; j < ListsObject[i].Count; j++)
                    {
                        arrRangeValue[i, j] = ListsObject[i][j];
                    }
                }
                return arrRangeValue;
            }
            return null;
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
                SpreadsheetsResource.ValuesResource.AppendRequest objAppendRequest = objSheetsService.Spreadsheets.Values.Append(objValueRange, Id, A1Range);
                objAppendRequest.ValueInputOption = SpreadsheetsResource.ValuesResource.AppendRequest.ValueInputOptionEnum.USERENTERED;
                objAppendRequest.Execute();
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
            if (Position >= 0)
            {
                Google.Apis.Sheets.v4.SheetsService objSheetsService = (Google.Apis.Sheets.v4.SheetsService)ActiveService;
                Google.Apis.Sheets.v4.Data.Request objRequest = new Request();
                Google.Apis.Sheets.v4.SpreadsheetsResource.GetRequest objGetRequest = objSheetsService.Spreadsheets.Get(Id);
                Google.Apis.Sheets.v4.Data.Spreadsheet objSpreadsheet = objGetRequest.Execute();
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
    }
}
