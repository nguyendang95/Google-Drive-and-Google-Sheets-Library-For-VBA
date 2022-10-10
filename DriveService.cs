using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using Google.Apis.Services;
using Google.Apis.Auth.OAuth2;
using System.Runtime.InteropServices;
using Google.Apis.Upload;
using GoogleApis;
using System.IO;
using Google;
using System.Linq.Expressions;
using System.Runtime.InteropServices.WindowsRuntime;
using System.Xml.Linq;

namespace GoogleDrive
{
    [Guid("6E234613-6381-472E-9343-58C3F0BB45CF")]
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    public interface IDriveService
    {
        string ClientId { get; set; }
        string ClientSecret { get; set; }
        string ApplicationName { get; set; }
        GDUploadStatus UploadStatus { get; set; }
        GDLoginStatus LoginStatus { get; set; }
        string LastUploadedFileName { get; set; }
        double BytesSent { get; set; }
        string LastUploadedFileId { get; set; }
        void UploadFileAsync(string FilePath, string FolderId = null);
        string UploadFile(string FilePath, string FolderId = null);
        GoogleDriveFile GetFileById(string FileId);
        string[] GetAllFileIdsInFolder(string FolderId = null);
        string CreateFolder(string FolderName, string DestinationFolderId = null);
        void LoginToGoogleDrive();
        GoogleDriveFile UpdateFile(string FileId, string NewFile);
        GoogleDriveFile CreateGoogleSheetsFile(string FileName, string FolderId = null);
        object ActiveService { get; set; }
    }
    public enum GDLoginStatus
    {
        gdLoginStatusNone,
        gdLoginStatusSuccess,
        gdLoginStatusError
    }
    public enum GDUploadStatus
    {
        gdUploadStatusNone,
        gdUploadStatusUploading,
        gdUploadStatusFailedFileNotFound,
        gdUploadStatusFailed,
        gdUploadStatusCompleted
    }
    [Guid("F7097191-5EA9-4564-A5C8-60D4787B01FE")]
    [ClassInterface(ClassInterfaceType.None)]
    public class DriveService : IDriveService
    {
        public string ClientId { get; set; }
        public string ClientSecret { get; set; }
        public string ApplicationName { get; set; }
        public GDUploadStatus UploadStatus { get; set; }
        public GDLoginStatus LoginStatus { get; set; }
        public string LastUploadedFileName { get; set; }
        public double BytesSent { get; set; }
        public string LastUploadedFileId { get; set; }
        public object ActiveService { get; set; }
        public DriveService()
        {
            UploadStatus = GDUploadStatus.gdUploadStatusNone;
            LoginStatus = GDLoginStatus.gdLoginStatusNone;
        }
        public void LoginToGoogleDrive()
        {
            try
            {
                string[] arrScope = { Google.Apis.Drive.v3.DriveService.Scope.Drive, Google.Apis.Drive.v3.DriveService.Scope.DriveFile };
                ClientSecrets objCS = new ClientSecrets
                {
                    ClientId = ClientId,
                    ClientSecret = ClientSecret
                };
                Google.Apis.Auth.OAuth2.UserCredential objUC = GoogleWebAuthorizationBroker.AuthorizeAsync(objCS, arrScope, System.Environment.UserName, System.Threading.CancellationToken.None, new Google.Apis.Util.Store.FileDataStore("GoogleDriveToken")).Result;
                Google.Apis.Drive.v3.DriveService objDS = new Google.Apis.Drive.v3.DriveService(new BaseClientService.Initializer
                {
                    HttpClientInitializer = objUC,
                    ApplicationName = ApplicationName
                });
                ActiveService = objDS;
                LoginStatus = GDLoginStatus.gdLoginStatusSuccess;
            }
            catch (Exception ex)
            {
                if (ex is AggregateException) { LoginStatus = GDLoginStatus.gdLoginStatusError; }
            }
        }
        public async void UploadFileAsync(string FilePath, string FolderId = null)
        {
            System.IO.FileInfo objFileInfo = new FileInfo(FilePath);
            if (!objFileInfo.Exists)
            {
                UploadStatus = GDUploadStatus.gdUploadStatusFailedFileNotFound;
                throw new FileNotFoundException("File not found");
            }
            Google.Apis.Drive.v3.Data.File objFile = new Google.Apis.Drive.v3.Data.File
            {
                Name = objFileInfo.Name,
                MimeType = GetMimeType(objFileInfo.FullName)
            };
            if (FolderId !=null)
            {
                string[] arrFolderName = new string[1];
                arrFolderName[0] = FolderId;
                objFile.Parents = ConvertArrayToList(arrFolderName);
            }
            try
            {
                System.IO.FileStream objFileStream = new System.IO.FileStream(FilePath, FileMode.Open);
                Google.Apis.Drive.v3.DriveService objDriveService = (Google.Apis.Drive.v3.DriveService)ActiveService;
                Google.Apis.Drive.v3.FilesResource.CreateMediaUpload objCreateMediaUpload = objDriveService.Files.Create(objFile, objFileStream, GetMimeType(objFileInfo.FullName));
                objCreateMediaUpload.Fields = "id";
                objCreateMediaUpload.SupportsAllDrives = true;
                objCreateMediaUpload.ProgressChanged += ObjCreateMediaUpload_ProgressChanged;
                objCreateMediaUpload.ResponseReceived += ObjCreateMediaUpload_ResponseReceived;
                UploadStatus = GDUploadStatus.gdUploadStatusUploading;
                await objCreateMediaUpload.UploadAsync();
                LastUploadedFileName = objFileInfo.FullName;
                objFileStream.Close();
                Google.Apis.Drive.v3.Data.File objUploadedFile = objCreateMediaUpload.ResponseBody;
                LastUploadedFileId = objUploadedFile.Id;
            }
            catch
            {
                UploadStatus = GDUploadStatus.gdUploadStatusFailed;
            }
        }
        public string UploadFile(string FilePath, string FolderId = null)
        {
            System.IO.FileInfo objFileInfo = new FileInfo(FilePath);
            if (!objFileInfo.Exists)
            {
                UploadStatus = GDUploadStatus.gdUploadStatusFailedFileNotFound;
                throw new FileNotFoundException("File not found");
            }
            Google.Apis.Drive.v3.Data.File objFile = new Google.Apis.Drive.v3.Data.File
            {
                Name = objFileInfo.Name,
                MimeType = GetMimeType(objFileInfo.FullName)
            };
            if (FolderId != null)
            {
                string[] arrFolderName = new string[1];
                arrFolderName[0] = FolderId;
                objFile.Parents = ConvertArrayToList(arrFolderName);
            }
            try
            {
                System.IO.FileStream objFileStream = new System.IO.FileStream(FilePath, FileMode.Open);
                Google.Apis.Drive.v3.DriveService objDriveService = (Google.Apis.Drive.v3.DriveService)ActiveService;
                Google.Apis.Drive.v3.FilesResource.CreateMediaUpload objCreateMediaUpload = objDriveService.Files.Create(objFile, objFileStream, GetMimeType(objFileInfo.FullName));
                objCreateMediaUpload.Fields = "id";
                objCreateMediaUpload.SupportsAllDrives = true;
                objCreateMediaUpload.ProgressChanged += ObjCreateMediaUpload_ProgressChanged1;
                objCreateMediaUpload.ResponseReceived += ObjCreateMediaUpload_ResponseReceived1;
                UploadStatus = GDUploadStatus.gdUploadStatusUploading;
                objCreateMediaUpload.Upload();
                LastUploadedFileName = objFileInfo.FullName;
                objFileStream.Close();
                Google.Apis.Drive.v3.Data.File objUploadedFile = objCreateMediaUpload.ResponseBody;
                LastUploadedFileId = objUploadedFile.Id;
                return objUploadedFile.Id;
            }
            catch
            {
                UploadStatus = GDUploadStatus.gdUploadStatusFailed;
            }
            return null;
        }

        private void ObjCreateMediaUpload_ResponseReceived1(Google.Apis.Drive.v3.Data.File obj)
        {
            UploadStatus = GDUploadStatus.gdUploadStatusCompleted;
        }

        private void ObjCreateMediaUpload_ProgressChanged1(IUploadProgress obj)
        {
            BytesSent = obj.BytesSent;
        }

        private void ObjCreateMediaUpload_ProgressChanged(IUploadProgress obj)
        {
            BytesSent = obj.BytesSent;
        }

        private void ObjCreateMediaUpload_ResponseReceived(Google.Apis.Drive.v3.Data.File obj)
        {
            UploadStatus = GDUploadStatus.gdUploadStatusCompleted;
        }

        private static string GetMimeType(string fileName)
        {
            string strMimeType = "application/unknown";
            string strExt = System.IO.Path.GetExtension(fileName).ToLower();
            Microsoft.Win32.RegistryKey objRegKey = Microsoft.Win32.Registry.ClassesRoot.OpenSubKey(strExt);
            if (objRegKey != null && objRegKey.GetValue("Content Type") != null)
                strMimeType = objRegKey.GetValue("Content Type").ToString();
            return strMimeType;
        }
        public string[] GetFolderIdByFolderName(string FolderName)
        {
            try
            {
                Google.Apis.Drive.v3.DriveService objDriveService = (Google.Apis.Drive.v3.DriveService)ActiveService;
                Google.Apis.Drive.v3.FilesResource.ListRequest objLR = objDriveService.Files.List();
                objLR.PageSize = 10;
                objLR.Q = "mimeType = 'application/vnd.google-apps.folder' and name = '" + FolderName + "'";
                objLR.Fields = "nextPageToken, files(id)";
                IList<Google.Apis.Drive.v3.Data.File> colFiles = objLR.Execute().Files;
                if (colFiles != null && colFiles.Count > 0)
                {
                    int i = 0;
                    string[] arrFolders = new string[colFiles.Count];
                    foreach (Google.Apis.Drive.v3.Data.File objFile in colFiles)
                    {
                        arrFolders[i] = objFile.Id;
                        i++;
                    }
                    return arrFolders;
                }
            }
            catch 
            {

            }
            return null;
        }
        public string[] GetAllFileIdsInFolder(string FolderId = null)
        {
            try
            {
                string strFolderId;
                if (FolderId == null) { strFolderId = "root"; }
                else { strFolderId = FolderId; }
                Google.Apis.Drive.v3.DriveService objDriveService = (Google.Apis.Drive.v3.DriveService)ActiveService;
                Google.Apis.Drive.v3.FilesResource.ListRequest objLR = objDriveService.Files.List();
                objLR.PageSize = 1000;
                objLR.Q = "'" + strFolderId + "' in parents";
                objLR.Fields = "nextPageToken, files(id)";
                IList<Google.Apis.Drive.v3.Data.File> colFiles = objLR.Execute().Files;
                if (colFiles != null && colFiles.Count > 0)
                {
                    int i = 0;
                    string[] arrFiles = new string[colFiles.Count];
                    foreach (Google.Apis.Drive.v3.Data.File objFile in colFiles)
                    {
                        arrFiles[i] = objFile.Id;
                        i++;
                    }
                    return arrFiles;
                }
            }
            catch 
            {
                return null;
            }
            return null;
        }
        public GoogleDriveFile GetFileById(string FileId)
        {
            Google.Apis.Drive.v3.DriveService objDriveService = (Google.Apis.Drive.v3.DriveService)ActiveService;
            Google.Apis.Drive.v3.FilesResource.GetRequest objRequest = objDriveService.Files.Get(FileId);
            objRequest.Fields = "*";
            Google.Apis.Drive.v3.Data.File objFile = objRequest.Execute();
            return InstantiateNewGoogleDriveFileObject(objFile);
        }
        private GoogleDriveFile InstantiateNewGoogleDriveFileObject(Google.Apis.Drive.v3.Data.File File)
        {
            if (File.MimeType != "application/vnd.google-apps.folder")
            {
                GoogleDriveFile objGoogleDriveFile = new GoogleDriveFile
                {
                    Name = File.Name,
                    WebViewLink = File.WebViewLink,
                    CreateTime = File.CreatedTime.ToString(),
                    Description = File.Description,
                    Id = File.Id,
                    MimeType = File.MimeType,
                    Parents = File.Parents.ToArray(),
                    Size = Convert.ToDouble(File.Size),
                    Version = Convert.ToDouble(File.Version),
                    ActiveService = ActiveService
                };
                return objGoogleDriveFile;
            }
            else
            {
                return null;
            }
        }
        private IList<string> ConvertArrayToList(string[] ArrayNames)
        {
            IList<string> List = ArrayNames.ToList();
            return List;
        }
        public string CreateFolder(string FolderName, string DestinationFolderId = null)
        {
            try
            {
                if (DestinationFolderId == null) { DestinationFolderId = "root"; }
                Google.Apis.Drive.v3.DriveService objDriveService = (Google.Apis.Drive.v3.DriveService)ActiveService;
                Google.Apis.Drive.v3.FilesResource.ListRequest objLR = objDriveService.Files.List();
                objLR.PageSize = 1000;
                objLR.Q = "'" + DestinationFolderId + "' in parents and mimeType = 'application/vnd.google-apps.folder' and name = '" + FolderName + "'";
                objLR.Fields = "nextPageToken, files(id)";
                IList<Google.Apis.Drive.v3.Data.File> colFiles = objLR.Execute().Files;
                if (colFiles != null || colFiles.Count == 0)
                {
                    Google.Apis.Drive.v3.Data.File objNewFolder = new Google.Apis.Drive.v3.Data.File
                    {
                        Name = FolderName,
                        MimeType = "application/vnd.google-apps.folder"
                    };
                    Google.Apis.Drive.v3.FilesResource.CreateRequest objCreateRequest = objDriveService.Files.Create(objNewFolder);
                    objCreateRequest.Fields = "id";
                    return objCreateRequest.Execute().Id;
                }
            }
            catch 
            {
                return null;
            }
            return null;
        }
        public GoogleDriveFile UpdateFile(string FileId, string NewFile)
        {
            try
            {
                Google.Apis.Drive.v3.DriveService objDriveService = (Google.Apis.Drive.v3.DriveService)ActiveService;
                System.IO.FileStream objFileStream = new System.IO.FileStream(NewFile, FileMode.Open);
                System.IO.FileInfo objFileInfo = new System.IO.FileInfo(NewFile);
                Google.Apis.Drive.v3.Data.File objFile = new Google.Apis.Drive.v3.Data.File
                {
                    Name = objFileInfo.Name
                };
                Google.Apis.Drive.v3.FilesResource.UpdateMediaUpload objUpdateRequest = objDriveService.Files.Update(objFile, FileId, objFileStream, GetMimeType(objFileInfo.Name));
                objUpdateRequest.Fields = "*";
                objUpdateRequest.Upload();
                objFileStream.Close();
                Google.Apis.Drive.v3.Data.File objNewFile = objUpdateRequest.ResponseBody;
                return InstantiateNewGoogleDriveFileObject(objNewFile);
            }
            catch 
            {
                return null;
            }
        }
        public GoogleDriveFile CreateGoogleSheetsFile(string FileName, string FolderId = null)
        {
            try
            {
                Google.Apis.Drive.v3.DriveService objDriveService = (Google.Apis.Drive.v3.DriveService)ActiveService;
                Google.Apis.Drive.v3.Data.File objNewSheet = new Google.Apis.Drive.v3.Data.File
                {
                    Name = FileName,
                    MimeType = "application/vnd.google-apps.spreadsheet",
                    Parents = new List<string> { FolderId}
                };
                Google.Apis.Drive.v3.FilesResource.CreateRequest objCreateRequest = objDriveService.Files.Create(objNewSheet);
                objCreateRequest.Fields = "*";
                Google.Apis.Drive.v3.Data.File objFile = objCreateRequest.Execute();
                return InstantiateNewGoogleDriveFileObject(objFile);
            }
            catch 
            {
                return null;
            }
        }
    }
}
