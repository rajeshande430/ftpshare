using System;
using System.Collections.Generic;
using System.IO;
using Microsoft.SharePoint.Client;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Security;
using System.Net;
using Microsoft.SharePoint.Client.Utilities;

namespace FTP_Share
{
    public static class SharepointHelper
    {
        private static List _List;
        private static ClientContext _ClientContext;
        private static Folder _Root;
        public static List<string> GetSubFolderNames()
        {
            return _Root.Folders.Select(_ => _.Name).ToList();
        }

        public static Folder ProjectFolder(string name)
        {
            return _Root.Folders.FirstOrDefault(_ => _.Name.Equals(name));
        }

        public static void Login(string username, string password, string urlsite, string libraryName = "Documents", string rootFolder = "PMRSFTP")
        {

            _ClientContext = new ClientContext(urlsite);
            var securePassword = new SecureString();
            foreach (char c in password.ToCharArray()) securePassword.AppendChar(c);

            _ClientContext.Credentials = new SharePointOnlineCredentials(username, securePassword);

            _List = _ClientContext.Web.Lists.GetByTitle(libraryName);
            _ClientContext.Load(_List);
            _ClientContext.ExecuteQuery();
            _ClientContext.Load(_List.RootFolder);
            _ClientContext.Load(_List.RootFolder.Folders);
            _ClientContext.ExecuteQuery();

            // navigate to the folder
            _Root = _List.RootFolder.Folders.First(_ => _.Name.Equals(rootFolder));
            _ClientContext.Load(_Root.Folders);
            _ClientContext.ExecuteQuery();



        }

        public static Folder GetSelectedSubFolder(string selectedSubFolder) => _Root.Folders.First(_ => _.Name.Equals(selectedSubFolder));

        public static Task UploadFolderToSharePoint(Item item, string projectName)
        {
            return Task.Run(() =>
            {
                UploadItemtoSharepoint(item, SharepointHelper.ProjectFolder(projectName), _ClientContext);
            });
        }

        public static string GetFolderHyperLink(Folder folder, int validDays = 14)
        {
            ClientContext ctx = _ClientContext;
            var expireDateforDownloadLink = DateTime.Now.AddDays(validDays).ToUniversalTime().ToString("yyyy-MM-ddTHH:mm:ssK");

            // Return the file object for the uploaded file.
            ctx.Load(folder.ListItemAllFields, item => item["EncodedAbsUrl"]);
            ctx.ExecuteQuery();

            var sharelink = Microsoft.SharePoint.Client.Web.CreateAnonymousLinkWithExpiration(ctx, folder.ListItemAllFields["EncodedAbsUrl"].ToString(), false, expireDateforDownloadLink);
            ctx.ExecuteQuery();
            var val = sharelink.Value;

            //return uploadFile.ListItemAllFields["EncodedAbsUrl"].ToString();
            return val = sharelink.Value;

        }

        public static bool IsFolderExistSP(string projectname, string foldername)
        {
            ClientContext ctx = _ClientContext;
            var projectFolder = ProjectFolder(projectname);

            ctx.Load(projectFolder.Folders);
            ctx.ExecuteQuery();

            return projectFolder.Folders.Any(_ => _.Name.Equals(foldername));

           
        }

        public static bool IsFileExistSP(ClientContext ctx, string selectedSubFolder, string filepath)
        {
            try
            {
                // Get the name of the file.
                string uniqueFileName = Path.GetFileName(filepath);
                var file = ctx.Web.GetFileByServerRelativeUrl(GetSelectedSubFolder(selectedSubFolder).ServerRelativeUrl + System.IO.Path.AltDirectorySeparatorChar + uniqueFileName);
                ctx.Load(file);
                ctx.ExecuteQuery();
                return true;
            }
            catch (Exception)
            {
                return false;
            }
        }

        public static void InsertEnquiryToSharepoint(string filepath, Folder selectedSubFolder, string sharedlink)
        {
            var filename = Path.GetFileName(filepath);
            var oList = _ClientContext.Web.Lists.GetByTitle("LOGS");
            ListItemCreationInformation itemCreateInfo = new ListItemCreationInformation();
            ListItem oListItem = oList.AddItem(itemCreateInfo);


            //Load the properties for the Web object.
            Web web = _ClientContext.Web;
            _ClientContext.Load(web);
            _ClientContext.ExecuteQuery();

            //Get the current user.
            _ClientContext.Load(web.CurrentUser);
            _ClientContext.ExecuteQuery();
            var currentUser = _ClientContext.Web.CurrentUser.Title;


            string versionNumber = "1.0";

            var uploadFile = _ClientContext.Web.GetFolderByServerRelativeUrl(selectedSubFolder.ServerRelativeUrl /*+ System.IO.Path.AltDirectorySeparatorChar + filename*/);
            _ClientContext.Load(uploadFile);
            //_ClientContext.Load(uploadFile.Versions);
            _ClientContext.ExecuteQuery();

            FileVersion version = null;
            //if (uploadFile.Versions.Any())
            //{

            //    version = uploadFile.Versions.LastOrDefault();
            //    versionNumber = (float.Parse(version.VersionLabel) + 1).ToString("#.0");
            //}



            oListItem["Title"] = currentUser;
            oListItem["Version_x0020_No"] = versionNumber;

            oListItem["File_x0020_Name"] = Path.GetFileName(filepath);

            oListItem["LocalPath"] = filepath;
            oListItem["Date"] = version?.Created ?? DateTime.Now;
            oListItem["FTP_x0020_Link"] = sharedlink;

            oListItem.Update();
            _ClientContext.ExecuteQuery();


        }

        public static bool CheckIfTheUserIsInTheNetwork()
        {

            IPHostEntry host;
            string localIP = "";

            host = Dns.GetHostEntry(Dns.GetHostName());

            foreach (IPAddress ip in host.AddressList)
            {
                if (ip.AddressFamily.ToString().Equals("InterNetwork"))
                {
                    localIP = ip.ToString();
                    var splitIPs = localIP.Split('.');
                    //192.168.6.1 - 192.168.6.255
                    if (splitIPs[0].Equals("192") && splitIPs[1].Equals("168") && splitIPs[2].Equals("6"))
                    {

                        return true;
                        //int last = int.Parse(splitIPs[3]);

                        //if(last >= 1 && last <= 255)
                        //{
                        //    return true;
                        //}
                    }

                }
            }


            return false;
        }

        public static string UploadFileSlicePerSlice(ClientContext ctx, Folder selectedSubFolder, string filepath, int fileChunkSizeInMB = 10)
        {
            // Each sliced upload requires a unique ID.
            Guid uploadId = Guid.NewGuid();

            // Get the name of the file.
            string uniqueFileName = Path.GetFileName(filepath);


            // Get the folder to upload into. 
            // List docs = ctx.Web.Lists.GetByTitle(libraryName);


            // File object.
            Microsoft.SharePoint.Client.File uploadFile;

            // Calculate block size in bytes.
            int blockSize = fileChunkSizeInMB * 1024 * 1024;

            var expireDateforDownloadLink = DateTime.Now.AddDays(14).ToUniversalTime().ToString("yyyy-MM-ddTHH:mm:ssK");


            // Get the size of the file.
            long fileSize = new FileInfo(filepath).Length;

            // If file already exist
            if (SharepointHelper.IsFileExistSP(ctx, selectedSubFolder.Name, filepath))
            {
                var result = System.Windows.MessageBox.Show($"file name '{Path.GetFileName(filepath)}' already exist. Do you wish to overwrite it?", "Overwrite Existing File", System.Windows.MessageBoxButton.YesNo, System.Windows.MessageBoxImage.Warning);

                if (result == System.Windows.MessageBoxResult.No)
                {
                    return string.Empty;
                }
            }

            if (fileSize <= blockSize)
            {
                // Use regular approach.
                using (FileStream fs = new FileStream(filepath, FileMode.Open))
                {

                    FileCreationInformation fileInfo = new FileCreationInformation();
                    fileInfo.ContentStream = fs;
                    fileInfo.Url = uniqueFileName;
                    fileInfo.Overwrite = true;
                    uploadFile = selectedSubFolder.Files.Add(fileInfo);
                    ctx.Load(uploadFile);
                    ctx.ExecuteQuery();

                    // Return the file object for the uploaded file.
                    ctx.Load(uploadFile.ListItemAllFields, item => item["EncodedAbsUrl"]);
                    ctx.ExecuteQuery();

                    var sharelink = Microsoft.SharePoint.Client.Web.CreateAnonymousLinkWithExpiration(ctx, uploadFile.ListItemAllFields["EncodedAbsUrl"].ToString(), false, expireDateforDownloadLink);
                    ctx.ExecuteQuery();
                    var val = sharelink.Value;

                    //return uploadFile.ListItemAllFields["EncodedAbsUrl"].ToString();
                    return val = sharelink.Value;
                }
            }
            else
            {
                // Use large file upload approach.
                ClientResult<long> bytesUploaded = null;

                FileStream fs = null;
                try
                {



                    fs = System.IO.File.Open(filepath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
                    using (BinaryReader br = new BinaryReader(fs))
                    {
                        byte[] buffer = new byte[blockSize];
                        Byte[] lastBuffer = null;
                        long fileoffset = 0;
                        long totalBytesRead = 0;
                        int bytesRead = 0;
                        bool first = true;
                        bool last = false;
                        string downloadablURL = "";



                        // Display the uploading progress 
                        ShareFTPForm.Object.txt_uploadingsize.Dispatcher.Invoke(System.Windows.Threading.DispatcherPriority.Normal,
                            new Action(() =>
                            { ShareFTPForm.Object.txt_uploadingsize.Visibility = System.Windows.Visibility.Visible; }));

                        // Read data from file system in blocks. 
                        while ((bytesRead = br.Read(buffer, 0, buffer.Length)) > 0)
                        {

                            // Display the uploading progress 
                            ShareFTPForm.Object.txt_uploadingsize.Dispatcher.Invoke(System.Windows.Threading.DispatcherPriority.Normal,
                                new Action(() =>
                                { ShareFTPForm.Object.txt_uploadingsize.Text = $"{totalBytesRead / 1024 / 1024} MB of {fileSize / 1024 / 1024} MB Uploaded"; }));

                            totalBytesRead = totalBytesRead + bytesRead;

                            // You've reached the end of the file.
                            if (totalBytesRead == fileSize)
                            {
                                last = true;
                                // Copy to a new buffer that has the correct size.
                                lastBuffer = new byte[bytesRead];
                                Array.Copy(buffer, 0, lastBuffer, 0, bytesRead);
                            }

                            if (first)
                            {
                                using (MemoryStream contentStream = new MemoryStream())
                                {
                                    // Add an empty file.
                                    FileCreationInformation fileInfo = new FileCreationInformation();
                                    fileInfo.ContentStream = contentStream;
                                    fileInfo.Url = uniqueFileName;
                                    fileInfo.Overwrite = true;
                                    uploadFile = selectedSubFolder.Files.Add(fileInfo);
                                    // Start upload by uploading the first slice. 
                                    using (MemoryStream s = new MemoryStream(buffer))
                                    {
                                        // Call the start upload method on the first slice.
                                        bytesUploaded = uploadFile.StartUpload(uploadId, s);
                                        ctx.ExecuteQuery();

                                        // Return the file object for the uploaded file.
                                        ctx.Load(uploadFile.ListItemAllFields, item => item["EncodedAbsUrl"]);
                                        ctx.ExecuteQuery();

                                        // Get the downloadable link
                                        //var sharelink = Microsoft.SharePoint.Client.Web.CreateAnonymousLink(_ClientContext, uploadFile.ListItemAllFields["EncodedAbsUrl"].ToString(), false);
                                        var sharelink = Microsoft.SharePoint.Client.Web.CreateAnonymousLinkWithExpiration(ctx, uploadFile.ListItemAllFields["EncodedAbsUrl"].ToString(), false, expireDateforDownloadLink);
                                        ctx.ExecuteQuery();
                                        var val = sharelink.Value;
                                        downloadablURL = sharelink.Value;

                                        // downloadablURL = uploadFile.ListItemAllFields["EncodedAbsUrl"].ToString();
                                        // fileoffset is the pointer where the next slice will be added.
                                        fileoffset = bytesUploaded.Value;
                                    }

                                    // You can only start the upload once.
                                    first = false;
                                }
                            }
                            else
                            {
                                // Get a reference to your file.
                                uploadFile = ctx.Web.GetFileByServerRelativeUrl(selectedSubFolder.ServerRelativeUrl + System.IO.Path.AltDirectorySeparatorChar + uniqueFileName);

                                if (last)
                                {
                                    // Is this the last slice of data?
                                    using (MemoryStream s = new MemoryStream(lastBuffer))
                                    {
                                        // End sliced upload by calling FinishUpload.
                                        uploadFile = uploadFile.FinishUpload(uploadId, fileoffset, s);
                                        ctx.ExecuteQuery();


                                        return downloadablURL;
                                    }
                                }
                                else
                                {
                                    using (MemoryStream s = new MemoryStream(buffer))
                                    {
                                        // Continue sliced upload.
                                        bytesUploaded = uploadFile.ContinueUpload(uploadId, fileoffset, s);
                                        ctx.ExecuteQuery();
                                        // Update fileoffset for the next slice.
                                        fileoffset = bytesUploaded.Value;
                                    }
                                }
                            }

                        } // while ((bytesRead = br.Read(buffer, 0, buffer.Length)) > 0)
                    }
                }
                finally
                {
                    if (fs != null)
                    {
                        fs.Dispose();
                    }
                }
            }

            return null;
        }

        public static Folder UploadFolder(Item item, Folder parentFolder, ClientContext clientContext)
        {
            var newfolder = parentFolder.Folders.Add(item.Name);
            clientContext.Load(newfolder);
            clientContext.ExecuteQuery();

            return newfolder;
        }

        public static void UploadFile(Item item, Folder parentFolder, ClientContext clientContext, int fileChunkSizeInMB = 10)
        {
            try
            {
                UploadFileSlicePerSlice(clientContext, parentFolder, item.FullPath);
            }
            catch (Exception ex)
            {
                System.Windows.MessageBox.Show("Something bad happend\n" + ex.Message, "Error", System.Windows.MessageBoxButton.OK, System.Windows.MessageBoxImage.Warning);
            }
        }


        public static void UploadItemtoSharepoint(Item item, Folder parentFolder, ClientContext clientContext)
        {
            var parent = UploadFolder(item, parentFolder, clientContext);
            item.Folder = parent;

            var files = item.Items.Where(_ => _.Type == ItemType.File);
            foreach (var file in files)
            {
                UploadFile(file, parent, clientContext);
            }


            var folders = item.Items.Where(_ => _.Type == ItemType.Folder);
            foreach (var folder in folders)
            {
                UploadItemtoSharepoint(folder, parent, clientContext);
            }

        }



    }
}
