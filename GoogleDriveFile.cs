using Google;
using GoogleDrive;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Runtime.Serialization;
using System.Security.Cryptography.X509Certificates;

namespace GoogleApis
{
    [Guid("C792B2AA-0ABA-4033-B492-2E8E16ECC8F2")]
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    public interface IGoogleDriveFile
    {
        string Name { get; set; }
        string Description { get; set; }
        string WebViewLink { get; set; }
        double Size { get; set; }
        string CreateTime { get; set; }
        string Id { get; set; }
        string[] Parents { get; set; }
        double Version { get; set; }
        string MimeType { get; set; }
        object ActiveService { get; set; }
        double BytesDownloaded { get; set; }
        GDDownloadStatus DownloadStatus { get; set; }
        GDDeleteFile DeleteFile(string FileId = null);
        GDMoveFile MoveFile(string FolderName, string FileId = null);
        void DownloadFileAsync(string DestinationFolder, string FileId = null);
        void DownloadFile(string DestinationFolder, string FileId = null);
        string CreateShareableLink(GDPermissionType Type, GDPermissionRole Role, string FileId = null, string EmailAddress = null, string ExpirationTime = null, string Domain = null, string DisplayName = null);
        GDRenameFile RenameFile(string NewName, string FileId = null);
        GDCopyFile CopyFile(string DestinationFolderId = null, string FileId = null);
        string[] GetListOfPermissionIds(string FileId = null);
        string CreatePermission(GDPermissionType Type, GDPermissionRole Role, string FileId = null, string EmailAddress = null, string ExpirationTime = null, string Domain = null, string DisplayName = null);
    }
    public enum GDPermissionType
    {
        gdPermissionTypeUser,
        gdPermissionTypeGroup,
        gdPermissionTypeDomain,
        gdPermissionTypeAnyone
    }
    public enum GDPermissionRole
    {
        gdPermissionRoleOwner,
        gdPermissionRoleOrganizer,
        gdPermissionRoleFileOrganizer,
        gdPermissionRoleWriter,
        gdPermissionRoleCommenter,
        gdPermissionRoleReader
    }
    public enum GDDeleteFile
    {
        gdDeleteFileCompleted,
        gdDeleteFileFailed
    }
    public enum GDMoveFile
    {
        gdMoveFileCompleted,
        gdMoveFileFailed
    }
    public enum GDDownloadStatus
    {
        gdDownloadStatusNone,
        gdDownloadStatusCompleted,
        gdDownloadStatusFailed,
        gdDownloadStatusDownloading
    }
    public enum GDCopyFile
    {
        gdCopyFileCompleted,
        gdCopyFileFailed
    }
    public enum GDRenameFile
    {
        gdRenameFileCompleted,
        gdRenameFileFailed
    }
    [Guid("6C84F219-EACB-44E3-BAF9-1A4221D724FD")]
    [ClassInterface(ClassInterfaceType.None)]
    public class GoogleDriveFile : IGoogleDriveFile
    {
        public string Name { get; set; }
        public string WebViewLink { get; set; }
        public double Size { get; set; }
        public string CreateTime { get; set; }
        public string Description { get; set; }
        public string Id { get; set; }
        public string[] Parents { get; set; }
        public double Version { get; set; }
        public string MimeType { get; set; }
        public double BytesDownloaded { get; set; }
        public object ActiveService { get; set; }
        public GDDownloadStatus DownloadStatus { get; set; }
        public GoogleDriveFile()
        {
            DownloadStatus = GDDownloadStatus.gdDownloadStatusNone;
        }
        public string CreatePermission(GDPermissionType Type, GDPermissionRole Role, string FileId = null, string EmailAddress = null, string ExpirationTime = null, string Domain = null, string DisplayName = null)
        {
            if (FileId == null) { FileId = Id; }
            Google.Apis.Drive.v3.DriveService GoogleDriveService = (Google.Apis.Drive.v3.DriveService)ActiveService;
            Google.Apis.Drive.v3.Data.Permission objPermission = new Google.Apis.Drive.v3.Data.Permission();
            switch (Type)
            {
                case GDPermissionType.gdPermissionTypeUser:
                    objPermission.Type = "user";
                    objPermission.DisplayName = DisplayName;
                    break;
                case GDPermissionType.gdPermissionTypeDomain:
                    objPermission.Domain = Domain;
                    objPermission.DisplayName = DisplayName;
                    break;
                case GDPermissionType.gdPermissionTypeGroup:
                    objPermission.Type = "user";
                    objPermission.DisplayName = DisplayName;
                    break;
                case GDPermissionType.gdPermissionTypeAnyone:
                    objPermission.Type = "anyone";
                    break;
            }
            switch (Role)
            {
                case GDPermissionRole.gdPermissionRoleCommenter:
                    objPermission.Role = "commenter";
                    break;
                case GDPermissionRole.gdPermissionRoleReader:
                    objPermission.Role = "reader";
                    break;
                case GDPermissionRole.gdPermissionRoleOwner:
                    objPermission.Role = "owner";
                    break;
                case GDPermissionRole.gdPermissionRoleWriter:
                    objPermission.Role = "writer";
                    break;
                case GDPermissionRole.gdPermissionRoleOrganizer:
                    objPermission.Role = "organizer";
                    break;
            }
            if (ExpirationTime != null)
            {
                if (Convert.ToDateTime(ExpirationTime) > DateTime.Now)
                {
                    objPermission.ExpirationTime = Convert.ToDateTime(ExpirationTime);
                }
                else
                {
                    throw new InvalidDateTimeValueException("Invalid datetime. The value must be greater current time.");
                }
            }
            if (Type == GDPermissionType.gdPermissionTypeUser || Type == GDPermissionType.gdPermissionTypeGroup || Type == GDPermissionType.gdPermissionTypeDomain)
            {
                if (EmailAddress != null)
                {
                    objPermission.EmailAddress = EmailAddress;
                }
                else
                {
                    throw new BlankEmailAddressOnUserGroupDomainException("For type 'user' and 'group', you must specify an email address. For type 'domain', you must specify domain.");
                }
            }
            try
            {
                return GoogleDriveService.Permissions.Create(objPermission, FileId).Execute().Id;
            }
            catch
            {
                return null;
            }
        }
        public string[] GetListOfPermissionIds(string FileId = null)
        {
            if (FileId == null) { FileId = Id; }
            Google.Apis.Drive.v3.DriveService objDriveService = (Google.Apis.Drive.v3.DriveService)ActiveService;
            Google.Apis.Drive.v3.FilesResource.GetRequest objGetRequest = objDriveService.Files.Get(FileId);
            objGetRequest.Fields = "*";
            IList<Google.Apis.Drive.v3.Data.Permission> PermissionList = objGetRequest.Execute().Permissions;
            string[] Result = new string[PermissionList.Count];
            int i = 0;
            foreach (Google.Apis.Drive.v3.Data.Permission objPermission in PermissionList)
            {
                Result[i] = objPermission.Id;
                i++;
            }
            return Result;
        }
        public string CreateShareableLink(GDPermissionType Type, GDPermissionRole Role, string FileId = null, string EmailAddress = null, string ExpirationTime = null, string Domain = null, string DisplayName = null)
        {
            if (FileId == null) { FileId = Id; }
            Google.Apis.Drive.v3.DriveService GoogleDriveService = (Google.Apis.Drive.v3.DriveService)ActiveService;
            Google.Apis.Drive.v3.Data.Permission objPermission = new Google.Apis.Drive.v3.Data.Permission();
            switch (Type)
            {
                case GDPermissionType.gdPermissionTypeUser:
                    objPermission.Type = "user";
                    objPermission.DisplayName = DisplayName;
                    break;
                case GDPermissionType.gdPermissionTypeDomain:
                    objPermission.Domain = Domain;
                    objPermission.DisplayName = DisplayName;
                    break;
                case GDPermissionType.gdPermissionTypeGroup:
                    objPermission.Type = "user";
                    objPermission.DisplayName = DisplayName;
                    break;
                case GDPermissionType.gdPermissionTypeAnyone:
                    objPermission.Type = "anyone";
                    break;
            }
            switch (Role)
            {
                case GDPermissionRole.gdPermissionRoleCommenter:
                    objPermission.Role = "commenter";
                    break;
                case GDPermissionRole.gdPermissionRoleReader:
                    objPermission.Role = "reader";
                    break;
                case GDPermissionRole.gdPermissionRoleOwner:
                    objPermission.Role = "owner";
                    break;
                case GDPermissionRole.gdPermissionRoleWriter:
                    objPermission.Role = "writer";
                    break;
                case GDPermissionRole.gdPermissionRoleOrganizer:
                    objPermission.Role = "organizer";
                    break;
            }
            if (ExpirationTime != null)
            {
                if (Convert.ToDateTime(ExpirationTime) > DateTime.Now)
                {
                    objPermission.ExpirationTime = Convert.ToDateTime(ExpirationTime);
                }
                else
                {
                    throw new InvalidDateTimeValueException("Invalid datetime. The value must be greater current time.");
                }
            }
            if (Type == GDPermissionType.gdPermissionTypeUser || Type == GDPermissionType.gdPermissionTypeGroup || Type == GDPermissionType.gdPermissionTypeDomain)
            {
                if (EmailAddress != null)
                {
                    objPermission.EmailAddress = EmailAddress;
                }
                else
                {
                    throw new BlankEmailAddressOnUserGroupDomainException("For type 'user' and 'group', you must specify an email address. For type 'domain', you must specify domain.");
                }
            }
            try
            {
                GoogleDriveService.Permissions.Create(objPermission, FileId).Execute();
                Google.Apis.Drive.v3.FilesResource.GetRequest objRequest = GoogleDriveService.Files.Get(FileId);
                objRequest.Fields = "webViewLink";
                return objRequest.Execute().WebViewLink;
            }
            catch
            {
                return null;
            }
        }
        public async void DownloadFileAsync(string DestinationFolder, string FileId = null)
        {
            try
            {
                DownloadStatus = GDDownloadStatus.gdDownloadStatusNone;
                if (FileId == null) { FileId = Id; }
                Google.Apis.Drive.v3.DriveService objDriveService = (Google.Apis.Drive.v3.DriveService)ActiveService;
                Google.Apis.Drive.v3.FilesResource.GetRequest objRequest = objDriveService.Files.Get(FileId);
                System.IO.FileStream objFileStream = new System.IO.FileStream(DestinationFolder, System.IO.FileMode.OpenOrCreate, System.IO.FileAccess.Write);
                objRequest.MediaDownloader.ProgressChanged += MediaDownloader_ProgressChanged;
                objRequest.SupportsAllDrives = true;
                await objRequest.DownloadAsync(objFileStream);
                objFileStream.Close();
            }
            catch
            {
                DownloadStatus = GDDownloadStatus.gdDownloadStatusFailed;
            }

        }
        private void MediaDownloader_ProgressChanged(Google.Apis.Download.IDownloadProgress obj)
        {
            switch (obj.Status)
            {
                case Google.Apis.Download.DownloadStatus.Failed:
                    DownloadStatus = GDDownloadStatus.gdDownloadStatusFailed;
                    throw new DownloadFileFailedException("Failed to download file. Please check your internet connection and try again!");
                case Google.Apis.Download.DownloadStatus.Downloading:
                    DownloadStatus = GDDownloadStatus.gdDownloadStatusDownloading;
                    BytesDownloaded = obj.BytesDownloaded;
                    break;
                case Google.Apis.Download.DownloadStatus.Completed:
                    DownloadStatus = GDDownloadStatus.gdDownloadStatusCompleted;
                    break;
            }
        }
        public void DownloadFile(string DestinationFolder, string FileId = null)
        {
            try
            {
                DownloadStatus = GDDownloadStatus.gdDownloadStatusNone;
                if (FileId == null) { FileId = Id; }
                Google.Apis.Drive.v3.DriveService objDriveService = (Google.Apis.Drive.v3.DriveService)ActiveService;
                Google.Apis.Drive.v3.FilesResource.GetRequest objRequest = objDriveService.Files.Get(FileId);
                System.IO.FileStream objFileStream = new System.IO.FileStream(DestinationFolder, System.IO.FileMode.OpenOrCreate, System.IO.FileAccess.Write);
                objRequest.MediaDownloader.ProgressChanged += MediaDownloader_ProgressChanged1;
                objRequest.SupportsAllDrives = true;
                objRequest.Download(objFileStream);
                objFileStream.Close();
            }
            catch
            {
                DownloadStatus = GDDownloadStatus.gdDownloadStatusFailed; 
            }
        }
        private void MediaDownloader_ProgressChanged1(Google.Apis.Download.IDownloadProgress obj)
        {
            switch (obj.Status)
            {
                case Google.Apis.Download.DownloadStatus.Failed:
                    DownloadStatus = GDDownloadStatus.gdDownloadStatusFailed;
                    throw new DownloadFileFailedException("Failed to download file. Please check your internet connection and try again!");
                case Google.Apis.Download.DownloadStatus.Downloading:
                    DownloadStatus = GDDownloadStatus.gdDownloadStatusDownloading;
                    BytesDownloaded = obj.BytesDownloaded;
                    break;
                case Google.Apis.Download.DownloadStatus.Completed:
                    DownloadStatus = GDDownloadStatus.gdDownloadStatusCompleted;
                    break;
            }
        }
        public GDDeleteFile DeleteFile(string FileId = null)
        {
            if (FileId == null) { FileId = Id; }
            try
            {
                Google.Apis.Drive.v3.DriveService objDriveService = (Google.Apis.Drive.v3.DriveService)ActiveService;
                Google.Apis.Drive.v3.FilesResource.DeleteRequest objRequest = objDriveService.Files.Delete(FileId);
                objRequest.SupportsAllDrives = true;
                objRequest.Execute();
                Name = null;
                WebViewLink = null;
                CreateTime = null;
                Description = null;
                Id = null;
                MimeType = null;
                Parents = null;
                Size = 0;
                Version = 0;
                return GDDeleteFile.gdDeleteFileCompleted;
            }
            catch
            {
                return GDDeleteFile.gdDeleteFileFailed;
            }
        }
        public GDMoveFile MoveFile(string FolderName, string FileId = null)
        {
            if (FileId == null) { FileId = Id; }
            try
            {
                Google.Apis.Drive.v3.DriveService objDriveService = (Google.Apis.Drive.v3.DriveService)ActiveService;
                Google.Apis.Drive.v3.FilesResource.GetRequest objGetRequest = objDriveService.Files.Get(FileId);
                objGetRequest.Fields = "*";
                Google.Apis.Drive.v3.Data.File objFile = objGetRequest.Execute();
                var arrPreviousParents = String.Join(",", objFile.Parents);
                Google.Apis.Drive.v3.FilesResource.UpdateRequest objUpdateRequest = objDriveService.Files.Update(new Google.Apis.Drive.v3.Data.File(), FileId);
                objUpdateRequest.Fields = "id, parents";
                objUpdateRequest.SupportsAllDrives = true;
                string strFolderId = GetFolderIdByFolderName(FolderName);
                objUpdateRequest.AddParents = strFolderId;
                objUpdateRequest.RemoveParents = arrPreviousParents;
                objFile = objUpdateRequest.Execute();
                UpdateFileObject(objFile);
                return GDMoveFile.gdMoveFileCompleted;
            }
            catch
            {
                return GDMoveFile.gdMoveFileFailed;
            }
        }
        private void UpdateFileObject(Google.Apis.Drive.v3.Data.File File)
        {
            Name = File.Name;
            WebViewLink = File.WebViewLink;
            CreateTime = File.CreatedTime.ToString();
            Description = File.Description;
            Id = File.Id;
            MimeType = File.MimeType;
            Parents = File.Parents.ToArray();
            Size = Convert.ToDouble(File.Size);
            Version = Convert.ToDouble(File.Version);
        }
        public GDRenameFile RenameFile(string NewName, string FileId = null)
        {
            try
            {
                if (FileId == null) { FileId = Id; }
                Google.Apis.Drive.v3.DriveService objDriveService = (Google.Apis.Drive.v3.DriveService)ActiveService;
                Google.Apis.Drive.v3.Data.File objFile = new Google.Apis.Drive.v3.Data.File
                {
                    Name = NewName
                };
                Google.Apis.Drive.v3.FilesResource.UpdateRequest objUpdateRequest = objDriveService.Files.Update(objFile, FileId);
                objUpdateRequest.Fields = "*";
                objUpdateRequest.SupportsAllDrives = true;
                Google.Apis.Drive.v3.Data.File objRenamedFile = objUpdateRequest.Execute();
                Name = objRenamedFile.Name;
                WebViewLink = objRenamedFile.WebViewLink;
                CreateTime = objRenamedFile.CreatedTime.ToString();
                Description = objRenamedFile.Description;
                Id = objRenamedFile.Id;
                MimeType = objRenamedFile.MimeType;
                Parents = objRenamedFile.Parents.ToArray();
                Size = Convert.ToDouble(objRenamedFile.Size);
                Version = Convert.ToDouble(objRenamedFile.Version);
                return GDRenameFile.gdRenameFileCompleted;
            }
            catch
            {
                return GDRenameFile.gdRenameFileFailed;
            }
        }
        private string GetFolderIdByFolderName(string FolderName)
        {
            try
            {
                Google.Apis.Drive.v3.DriveService objDriveService = (Google.Apis.Drive.v3.DriveService)ActiveService;
                Google.Apis.Drive.v3.FilesResource.ListRequest objLR = objDriveService.Files.List();
                objLR.PageSize = 10;
                objLR.Q = "mimeType = 'application/vnd.google-apps.folder' and name = '" + FolderName + "'";
                objLR.Fields = "nextPageToken, files(id, name)";
                IList<Google.Apis.Drive.v3.Data.File> colFiles = objLR.Execute().Files;
                if (colFiles != null && colFiles.Count > 0)
                {
                    foreach (Google.Apis.Drive.v3.Data.File objFile in colFiles)
                    {
                        return objFile.Id;
                    }
                }
            }
            catch
            {
                return null;
            }
            return null;
        }
        public GDCopyFile CopyFile(string DestinationFolderId = null, string FileId = null)
        {
            if (FileId == null) { FileId = Id; }
            try
            {
                Google.Apis.Drive.v3.DriveService objDriveService = (Google.Apis.Drive.v3.DriveService)ActiveService;
                Google.Apis.Drive.v3.Data.File objCopiedFile = new Google.Apis.Drive.v3.Data.File
                {
                    Parents = new List<string> { DestinationFolderId }
                };
                Google.Apis.Drive.v3.FilesResource.CopyRequest objCopyRequest = objDriveService.Files.Copy(objCopiedFile, FileId);
                objCopyRequest.Fields = "*";
                objCopyRequest.SupportsAllDrives = true;
                Google.Apis.Drive.v3.Data.File objNewFile = objCopyRequest.Execute();
                UpdateFileObject(objNewFile);
                return GDCopyFile.gdCopyFileCompleted;
            }
            catch
            {
                return GDCopyFile.gdCopyFileFailed;
            }
        }
    }
}
