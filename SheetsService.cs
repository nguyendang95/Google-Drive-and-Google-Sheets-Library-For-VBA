using Google.Apis.Auth.OAuth2;
using Google.Apis.Services;
using GoogleDrive;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;

namespace GoogleApis
{
    [Guid("C792B2AA-0ABA-4033-B492-2E8E16ECC8F3")]
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    public interface ISheetsService
    {
        string ClientId { get; set; }
        string ClientSecret { get; set; }
        string ApplicationName { get; set; }
        GSLoginStatus LoginStatus { get; set; }
        object ActiveService { get; set; }
        GSLoginStatus LoginToGoogleSheets();
        GoogleSheetsFile OpenSheetsFile(string SheetsFileId);
    }
    public enum GSLoginStatus
    {
        gsLoginStatusNone,
        gsLoginStatusCompleted,
        gsLoginStatusFailed
    }
    [Guid("6C84F220-EACB-44E3-BAF9-1A4221D724FE")]
    [ClassInterface(ClassInterfaceType.None)]
    public class SheetsService : ISheetsService
    {
        public string ClientId { get; set; }
        public string ClientSecret { get; set; }
        public string ApplicationName { get; set; }
        public object ActiveService { get; set; }
        public GSLoginStatus LoginStatus { get; set; }
        public SheetsService()
        {
            LoginStatus = GSLoginStatus.gsLoginStatusNone;
        }
        public GSLoginStatus LoginToGoogleSheets()
        {
            try
            {
                string[] arrScope = { Google.Apis.Sheets.v4.SheetsService.Scope.Spreadsheets };
                ClientSecrets objCS = new ClientSecrets
                {
                    ClientId = ClientId,
                    ClientSecret = ClientSecret
                };
                Google.Apis.Auth.OAuth2.UserCredential objUC = GoogleWebAuthorizationBroker.AuthorizeAsync(objCS, arrScope, System.Environment.UserName, System.Threading.CancellationToken.None, new Google.Apis.Util.Store.FileDataStore("GoogleSheetsToken")).Result;
                Google.Apis.Sheets.v4.SheetsService objSheetsService = new Google.Apis.Sheets.v4.SheetsService(new BaseClientService.Initializer
                {
                    HttpClientInitializer = objUC,
                    ApplicationName = ApplicationName
                });
                ActiveService = objSheetsService;
                LoginStatus = GSLoginStatus.gsLoginStatusCompleted;
                return LoginStatus;
            }
            catch (Exception ex)
            {
                if (ex is AggregateException) { LoginStatus = GSLoginStatus.gsLoginStatusFailed; }
            }
            return LoginStatus;
        }
        public GoogleSheetsFile OpenSheetsFile(string SheetsFileId)
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
    }
}
