using System;
using Google.Apis.Auth.OAuth2;
using Google.Apis.Download;
using Google.Apis.Drive.v3;
using Google.Apis.Services;
using System.Collections.Generic;
using System.IO;
using System.Security.Cryptography.X509Certificates;
using System.Configuration;
using System.Threading;
using Google.Apis.Util.Store;
//using log4net;
using System.Linq;
using System.Web;
using Google.Apis.Requests;
using Google.Apis.Drive.v3.Data;
using System.Web.Hosting;

namespace GoogleServices
{
    public static class GDriveUtils
    {
       // private static readonly ILog log = LogManager.GetLogger(typeof(GDriveUtils));

        /// <summary>
        /// Atributo que tendra la conexion con la API de Google, se invoca una sola vez para uso global
        /// </summary>
        public static DriveService service;

        static GDriveUtils()
        {
            service = AuthenticateServiceAccount();
        }


        /// <summary>
        /// Realiza la autenticacion para conectar con la API de Google Drive
        /// </summary>
        /// <returns>DriveService con el acceso listo</returns>
        public static DriveService AuthenticateServiceAccount()
        {

            try
            {
                string serviceAccountCredentialFilePath = getCredentialFilePath();
                //log.Debug("serviceAccountCredentialFilePath:" + getCredentialFilePath());
                if (string.IsNullOrEmpty(serviceAccountCredentialFilePath))
                    throw new Exception("Path to the service account credentials file is required.");

                if (!System.IO.File.Exists(serviceAccountCredentialFilePath))
                    throw new Exception("The service account credentials file does not exist at: " + serviceAccountCredentialFilePath);

                if (string.IsNullOrEmpty(ConfigurationManager.AppSettings["GDRIVE_CRED_EMAIL"]))
                    throw new Exception("ServiceAccountEmail is required.");

                //log.Debug("serviceAccountCredentialFilePath:" + serviceAccountCredentialFilePath);
                //log.Debug("AppDomain.CurrentDomain.BaseDirectory:" + AppDomain.CurrentDomain.BaseDirectory);
                // These are the scopes of permissions you need. It is best to request only what you need and not all of them
                string[] scopes = { DriveService.Scope.Drive };

                // For Json file
                if (Path.GetExtension(serviceAccountCredentialFilePath).ToLower() == ".json")
                {
                    //GoogleCredential credential;
                    String FilePath = AppDomain.CurrentDomain.BaseDirectory + "\\GoogleAPI\\DriveCredentials";
                    UserCredential credential;
                    using (var stream = new FileStream(serviceAccountCredentialFilePath, FileMode.Open, FileAccess.Read))
                    {
                        //credential = GoogleCredential.FromStream(stream).CreateScoped(scopes);
                        credential = GoogleWebAuthorizationBroker.AuthorizeAsync(
                            GoogleClientSecrets.Load(stream).Secrets,
                            scopes,
                            "user",
                            CancellationToken.None,
                            new FileDataStore(FilePath, true)).Result;

                    }

                    // Create the  Analytics service.
                    return new DriveService(new BaseClientService.Initializer()
                    {
                        HttpClientInitializer = credential,
                        ApplicationName = "OGEI APPS",
                    });
                }
                else if (Path.GetExtension(serviceAccountCredentialFilePath).ToLower() == ".p12")
                {   // If its a P12 file

                    var certificate = new X509Certificate2(serviceAccountCredentialFilePath, "notasecret", X509KeyStorageFlags.MachineKeySet | X509KeyStorageFlags.Exportable);
                    var credential = new ServiceAccountCredential(new ServiceAccountCredential.Initializer(ConfigurationManager.AppSettings["GDRIVE_CRED_EMAIL"])
                    {
                        Scopes = scopes
                    }.FromCertificate(certificate));

                    // Create the  Drive service.
                    return new DriveService(new BaseClientService.Initializer()
                    {
                        HttpClientInitializer = credential,
                        ApplicationName = "Drive Authentication Sample",
                    });
                }
                else
                {
                    throw new Exception("Unsupported Service accounts credentials.");
                }

            }
            catch (Exception ex)
            {
                //log.Error(String.Format("Error de comunicación: [{0}: {1}]\r\nStack Trace:\r\n{2}",
                //     ex.Source, ex.Message, ex.StackTrace));
                //if (ex.InnerException != null)
                //{
                //    log.Error(String.Format("Inner Exception: [{0}: {1}]\r\nStack Trace:\r\n{2}",
                //        ex.InnerException.Source, ex.InnerException.Message, ex.InnerException.StackTrace));
                //}
                throw new Exception("CreateServiceAccountDriveFailed", ex);
            }
        }

        /// <summary>
        /// Devuelve la dirección base del archivo json o p12 que se usara como credencial para acceder a la API
        /// </summary>
        /// <returns>String - dirección de la credencial</returns>
        public static string getCredentialFilePath()
        {
            //return AppDomain.CurrentDomain.BaseDirectory + "GoogleAPI" + "\\"+"DriveCredentials"+ "\\" + ConfigurationManager.AppSettings["GDRIVE_CRED_FILE_NAME"];
            return HostingEnvironment.MapPath(@"~/Statics/Google") + "\\" + ConfigurationManager.AppSettings["GDRIVE_CRED_FILE_NAME"];
        }

        /// <summary>
        /// 
        /// </summary>
        /// <returns></returns>
        public static List<GoogleDriveFiles> GetDriveFiles()
        {
            DriveService service = GDriveUtils.AuthenticateServiceAccount();

            // Define parameters of request.
            FilesResource.ListRequest FileListRequest = service.Files.List();
            FileListRequest.Fields = "nextPageToken, files(id, name, size, version, trashed, createdTime)";

            // List files.
            IList<Google.Apis.Drive.v3.Data.File> files = FileListRequest.Execute().Files;
            List<GoogleDriveFiles> FileList = new List<GoogleDriveFiles>();

            if (files != null && files.Count > 0)
            {
                foreach (var file in files)
                {
                    GoogleDriveFiles File = new GoogleDriveFiles
                    {
                        Id = file.Id,
                        Name = file.Name,
                        Size = file.Size,
                        Version = file.Version,
                        CreatedTime = file.CreatedTime
                    };
                    FileList.Add(File);
                }
            }
            return FileList;
        }

        /// <summary>
        /// Metodo que realiza un upload al repositorio drive, obtengo un archivo proveniente de POST (HttpPostedFileBase)
        /// </summary>
        /// <param name="file"></param>
        /// <returns>-1 si hubo un error en el upload, id del archivo drive en caso que el proceso haya sido exitoso</returns>
        public static string FileUpload(HttpPostedFileBase file, string carpeta = null)
        {
            try
            {

                string pathTemp = HttpContext.Current.Server.MapPath("~/GoogleDriveFilesTemp");

                if (!System.IO.Directory.Exists(pathTemp))
                    System.IO.Directory.CreateDirectory(pathTemp);

                if (file != null && file.ContentLength > 0)
                {
                    string path = Path.Combine(pathTemp,
                    Path.GetFileName(file.FileName));
                    file.SaveAs(path);

                    var FileMetaData = new Google.Apis.Drive.v3.Data.File();
                    FileMetaData.Name = Path.GetFileName(file.FileName);
                    FileMetaData.MimeType = MimeTypes.GetMimeType(path);
                    FileMetaData.Parents = new List<string>();
                    //FileMetaData.Parents.Add(ConfigurationManager.AppSettings["GDRIVE_SITRAD_FILE_ID"]);

                    //string carpeta = string.IsNullOrWhiteSpace(Request["carpeta"]) ? ConfigurationManager.AppSettings["GDRIVE_DEFAULT_FOLDER_ID"].ToString() : Request["carpeta"].ToString();

                    if (carpeta != null)  FileMetaData.Parents.Add(carpeta);

                    FilesResource.CreateMediaUpload request;
                    using (var stream = new System.IO.FileStream(path, System.IO.FileMode.Open))
                    {
                        request = service.Files.Create(FileMetaData, stream, FileMetaData.MimeType);
                        request.Fields = "id";
                        request.Upload();
                    }

                    //return request.ResponseBody.Id + "____" + FileMetaData.Name;
                    return request.ResponseBody.Id;
                }
            }
            catch (Exception ex)
            {
                //log.Error(String.Format("Error de comunicación: [{0}: {1}]\r\nStack Trace:\r\n{2}",
                //     ex.Source, ex.Message, ex.StackTrace));
                //if (ex.InnerException != null)
                //{
                //    log.Error(String.Format("Inner Exception: [{0}: {1}]\r\nStack Trace:\r\n{2}",
                //        ex.InnerException.Source, ex.InnerException.Message, ex.InnerException.StackTrace));
                //}
                return "-1";
            }
            return "-1";
        }


        /// <summary>
        /// Realiza upload de archivos de un path donde ya exista un archivo ubicado en un directorio temporal
        /// </summary>
        /// <param name="path"></param>
        /// <param name="fileName"></param>
        /// <returns></returns>
        public static string FileUpload(string path, string fileName, string parentFolder = null)
        {
            try
            {
                var folderId = "";
                //creando carpeta de resultar parentFolder != null
                if (parentFolder != null)
                {
                    FilesResource.ListRequest FileListRequest = service.Files.List();
                    FileListRequest.Fields = "nextPageToken, files(id, name, size, version, trashed, createdTime)";
                    IList<Google.Apis.Drive.v3.Data.File> files = FileListRequest.Execute().Files;


                    if (files != null && files.Count > 0)
                    {
                        foreach (var file in files)
                        {
                            GoogleDriveFiles File = new GoogleDriveFiles
                            {
                                Id = file.Id,
                                Name = file.Name,
                                Size = file.Size,
                                Version = file.Version,
                                CreatedTime = file.CreatedTime
                            };

                            if (file.Name == parentFolder && file.Trashed == false)
                            {
                                folderId = file.Id;
                                break;
                            }
                        }
                    }

                    if (folderId == "")
                    {

                        var folderMetadata = new Google.Apis.Drive.v3.Data.File()
                        {
                            Name = parentFolder,
                            MimeType = "application/vnd.google-apps.folder",
                            Parents = new List<string>
                            {
                                ConfigurationManager.AppSettings["GDRIVE_SITRAD_FILE_ID"]
                            }
                        };
                        var requestFolder = service.Files.Create(folderMetadata);
                        requestFolder.Fields = "id";
                        var folderRes = requestFolder.Execute();
                        folderId = folderRes.Id;
                    }
                }
                else folderId = ConfigurationManager.AppSettings["GDRIVE_SITRAD_FILE_ID"];

                var FileMetaData = new Google.Apis.Drive.v3.Data.File();
                FileMetaData.Name = fileName;
                FileMetaData.MimeType = MimeTypes.GetMimeType(fileName);
                FileMetaData.Parents = new List<string>();
                FileMetaData.Parents.Add(folderId);

                FilesResource.CreateMediaUpload request;
                using (var stream = new System.IO.FileStream(path, System.IO.FileMode.Open))
                {
                    request = service.Files.Create(FileMetaData, stream, FileMetaData.MimeType);
                    request.Fields = "id";
                    request.Upload();
                }


                /////////////////////////
                //var batch = new BatchRequest(service);
                //BatchRequest.OnResponse<Permission> callback = delegate(
                //    Permission permission,
                //    RequestError error,
                //    int index,
                //    System.Net.Http.HttpResponseMessage message)
                //{
                //    if (error != null)
                //    {
                //        // Handle error
                //        Console.WriteLine(error.Message);
                //    }
                //    else
                //    {
                //        Console.WriteLine("Permission ID: " + permission.Id);
                //    }
                //};

                //Permission userPermission = new Permission()
                //{
                //    Type = "anyone",
                //    Role = "owner",
                //    EmailAddress = "ogei_apps@vivienda.gob.pe",
                //    AllowFileDiscovery = true
                //};

                //var requestPerm = service.Permissions.Create(userPermission, request.ResponseBody.Id);
                //request.Fields = "id";
                //batch.Queue(requestPerm, callback);
                //var task = batch.ExecuteAsync();
                //////////////////////////

                return request.ResponseBody.Id;

            }
            catch (Exception ex)
            {
                //log.Error(String.Format("Error de comunicación: [{0}: {1}]\r\nStack Trace:\r\n{2}",
                //     ex.Source, ex.Message, ex.StackTrace));
                //if (ex.InnerException != null)
                //{
                //    log.Error(String.Format("Inner Exception: [{0}: {1}]\r\nStack Trace:\r\n{2}",
                //        ex.InnerException.Source, ex.InnerException.Message, ex.InnerException.StackTrace));
                //}
                return "-1";
            }
        }


        /// <summary>
        /// Realiza un download bajo conexion con la API
        /// </summary>
        /// <param name="fileId"></param>
        /// <returns>string FilePath</returns>
        public static string DownloadGoogleFile(string fileId)
        {
            //DriveService service = AuthenticateServiceAccount();
            string FolderPath = HttpContext.Current.Server.MapPath("~/GoogleDriveFilesTemp");
            FilesResource.GetRequest request = service.Files.Get(fileId);

            string FileName = request.Execute().Name;
            string FilePath = System.IO.Path.Combine(FolderPath, FileName);

            MemoryStream stream1 = new MemoryStream();

            request.MediaDownloader.ProgressChanged += (Google.Apis.Download.IDownloadProgress progress) =>
            {
                switch (progress.Status)
                {
                    case DownloadStatus.Downloading:
                        {
                            //Console.WriteLine(progress.BytesDownloaded);
                            break;
                        }
                    case DownloadStatus.Completed:
                        {
                            //Console.WriteLine("Download complete.");
                            SaveStream(stream1, FilePath);
                            break;
                        }
                    case DownloadStatus.Failed:
                        {
                            //Console.WriteLine("Download failed.");
                            break;
                        }
                }
            };
            request.Download(stream1);
            return FilePath;
        }



        private static void SaveStream(MemoryStream stream, string FilePath)
        {
            using (System.IO.FileStream file = new FileStream(FilePath, FileMode.Create, FileAccess.ReadWrite))
            {
                stream.WriteTo(file);
            }
        }


        /// <summary>
        /// Realiza movida de un archivo hacia otro directorio
        /// </summary>
        /// <param name="fileId"></param>
        /// <param name="folderId"></param>
        /// <returns></returns>
        public static string MoveFiles(String fileId, String folderId)
        {

            // Retrieve the existing parents to remove
            Google.Apis.Drive.v3.FilesResource.GetRequest getRequest = service.Files.Get(fileId);
            getRequest.Fields = "parents";
            Google.Apis.Drive.v3.Data.File file = getRequest.Execute();
            string previousParents = String.Join(",", file.Parents);

            // Move the file to the new folder
            Google.Apis.Drive.v3.FilesResource.UpdateRequest updateRequest = service.Files.Update(new Google.Apis.Drive.v3.Data.File(), fileId);
            updateRequest.Fields = "id, parents";
            updateRequest.AddParents = folderId;
            updateRequest.RemoveParents = previousParents;

            file = updateRequest.Execute();
            if (file != null)
            {
                return "Success";
            }
            else
            {
                return "Fail";
            }
        }

        /// <summary>
        /// Realiza copy/paste de un archivo hacia otro directorio
        /// </summary>
        /// <param name="fileId"></param>
        /// <param name="folderId"></param>
        /// <returns></returns>
        public static string CopyFiles(String fileId, String folderId)
        {
            // Retrieve the existing parents to remove
            Google.Apis.Drive.v3.FilesResource.GetRequest getRequest = service.Files.Get(fileId);
            getRequest.Fields = "parents";
            Google.Apis.Drive.v3.Data.File file = getRequest.Execute();

            // Copy the file to the new folder
            Google.Apis.Drive.v3.FilesResource.UpdateRequest updateRequest = service.Files.Update(new Google.Apis.Drive.v3.Data.File(), fileId);
            updateRequest.Fields = "id, parents";
            updateRequest.AddParents = folderId;
            file = updateRequest.Execute();
            if (file != null)
            {
                return "Success";
            }
            else
            {
                return "Fail";
            }
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="file_ID"></param>
        /// <returns></returns>
        public static string Delete(string file_ID)
        {

            try
            {
                //DriveService service = AuthenticateServiceAccount();
                // Initial validation.
                if (service == null)
                    throw new ArgumentNullException("service");

                if (file_ID == null)
                    throw new ArgumentNullException(file_ID);

                // Make the request.
                return service.Files.Delete(file_ID).Execute();
            }
            catch (Exception ex)
            {
                throw new Exception("Request Files.Delete failed.", ex);
                return "-1";
            }

            return "-1";
        }


        /// <summary>
        /// Elimina un archivo del repo de google drive
        /// </summary>
        /// <param name="files"></param>
        public static void DeleteFile(GoogleDriveFiles file)
        {
            try
            {
                // Initial validation.
                if (service == null)
                    throw new ArgumentNullException("service");

                if (file == null)
                    throw new ArgumentNullException(file.Id);

                // Make the request.
                service.Files.Delete(file.Id).Execute();
            }
            catch (Exception ex)
            {
                //log.Error(String.Format("Error de comunicación: [{0}: {1}]\r\nStack Trace:\r\n{2}",
                //     ex.Source, ex.Message, ex.StackTrace));
                //if (ex.InnerException != null)
                //{
                //    log.Error(String.Format("Inner Exception: [{0}: {1}]\r\nStack Trace:\r\n{2}",
                //        ex.InnerException.Source, ex.InnerException.Message, ex.InnerException.StackTrace));
                //}
                throw new Exception("Request Files.Delete failed.", ex);
            }
        }


    }

    /// <summary>
    /// Modelo que se usara para recibir los archivos de google drive
    /// </summary>
    public class GoogleDriveFiles
    {
        public string Id { get; set; }
        public string Name { get; set; }
        public long? Size { get; set; }
        public long? Version { get; set; }
        public DateTime? CreatedTime { get; set; }
    }

}

