using System;
using System.Configuration;
using System.Security;
using System.IO;
using Microsoft.SharePoint.Client;


namespace MSBuild2018
{
    class SharePointOnlineHelper
    {

        /// <summary>
        /// CreateFolderForEnvelope - Create a new folder for DocuSign envelope using SharePoint Online CSOM
        /// </summary>
        /// <param name="folderName"></param>
        public static void CreateFolderForEnvelope(string folderName)
        {
            string userName = ConfigurationManager.AppSettings["UserName"];
            string password = ConfigurationManager.AppSettings["Password"];
            string siteUrl = ConfigurationManager.AppSettings["SiteUrl"];
            string targetLibrary = ConfigurationManager.AppSettings["TargetLibrary"];

            using (ClientContext context = new ClientContext(siteUrl))
            {
                bool isOffice365 = false;
                Boolean.TryParse(ConfigurationManager.AppSettings["IsOffice365"], out isOffice365);
                if (isOffice365)
                {
                    SecureString securePassword = new SecureString();
                    foreach (char c in password.ToCharArray()) securePassword.AppendChar(c);

                    SharePointOnlineCredentials onlineCredentials = new SharePointOnlineCredentials(userName, securePassword);
                    context.Credentials = onlineCredentials;
                }
                else
                {
                    // OnPrem environment
                    return;
                }

                Web web = context.Web;
                context.Load(web);
                context.ExecuteQuery();

                Microsoft.SharePoint.Client.List list = context.Web.Lists.GetByTitle(targetLibrary);
                context.Load(list.RootFolder);
                context.ExecuteQuery();

                // create the new folder
                CreateFolder(web, list, list.RootFolder.ServerRelativeUrl, folderName);
            }
        }

        /// <summary>
        /// CreateFolder - create a folder in a document library using SharePoint Online CSOM
        /// </summary>
        /// <param name="web"></param>
        /// <param name="list"></param>
        /// <param name="folderRelativePath"></param>
        /// <param name="folderName"></param>
        /// <returns></returns>
        private static Microsoft.SharePoint.Client.Folder CreateFolder(Web web, Microsoft.SharePoint.Client.List list, string folderRelativePath, string folderName)
        {
            Microsoft.SharePoint.Client.Folder currentFolder = null;

            try
            {
                list.EnableFolderCreation = true;
                list.Update();
                web.Context.ExecuteQuery();

                ListItemCreationInformation itemCreateInfo = new ListItemCreationInformation();
                itemCreateInfo.UnderlyingObjectType = FileSystemObjectType.Folder;
                itemCreateInfo.LeafName = folderName;

                Microsoft.SharePoint.Client.ListItem newFolder = list.AddItem(itemCreateInfo);
                newFolder["Title"] = folderName;
                newFolder.Update();
                web.Context.ExecuteQuery();

                Microsoft.SharePoint.Client.Folder rootFolder = list.RootFolder;
                web.Context.Load(rootFolder);
                web.Context.ExecuteQuery();

                FolderCollection folders = rootFolder.Folders;
                web.Context.Load(folders);
                web.Context.ExecuteQuery();

                currentFolder = list.RootFolder.Folders[list.RootFolder.Folders.Count - 1];

            }
            catch (Exception ex)
            {
                Console.WriteLine("Error creating new folder named {0}", folderName);
                Console.WriteLine("Exception: {0}", ex.Message);
                return null;
            }

            return currentFolder;
        }

        /// <summary>
        /// UploadFileToSharePoint - upload file to SharePoint using SharePoint Online CSOM
        /// </summary>
        /// <param name="siteUrl"></param>
        /// <param name="userName"></param>
        /// <param name="password"></param>
        /// <param name="targetLibrary"></param>
        /// <param name="filePath"></param>
        public static void UploadFileToSharePoint(string siteUrl, string userName, string password, string targetLibrary, string filePath)
        {
            using (ClientContext context = new ClientContext(siteUrl))
            {
                Console.WriteLine("1. Connecting to SharePoint Online....");

                bool isOffice365 = false;
                Boolean.TryParse(ConfigurationManager.AppSettings["IsOffice365"], out isOffice365);
                if (isOffice365)
                {
                    SecureString securePassword = new SecureString();
                    foreach (char c in password.ToCharArray()) securePassword.AppendChar(c);

                    SharePointOnlineCredentials onlineCredentials = new SharePointOnlineCredentials(userName, securePassword);
                    context.Credentials = onlineCredentials;
                }
                else
                {
                    // this demo is only for SPO
                }

                Microsoft.SharePoint.Client.List list = context.Web.Lists.GetByTitle(targetLibrary);
                context.Load(list.RootFolder);
                context.ExecuteQuery();
                Console.WriteLine("2. Uploading {0} to SharePoint Online {1}....", filePath, siteUrl + @"/" + targetLibrary);

                string fileName = Path.GetFileName(filePath);
                string fileUrl = string.Format("{0}/{1}", list.RootFolder.ServerRelativeUrl, fileName);
                using (FileStream fs = new FileStream(filePath, FileMode.Open))
                {
                    Microsoft.SharePoint.Client.File.SaveBinaryDirect(context, fileUrl, fs, true);
                }
            }
        }

        /// <summary>
        /// UploadFileStreamToSharePoint - upload a MEmoryStream document to SharePoint using SharePoint Online CSOM
        /// </summary>
        /// <param name="siteUrl"></param>
        /// <param name="userName"></param>
        /// <param name="password"></param>
        /// <param name="targetLibrary"></param>
        /// <param name="file"></param>
        /// <param name="fileName"></param>
        /// <param name="envelopeID"></param>
        /// <param name="createNewFolder"></param>
        public static bool UploadFileStreamToSharePoint(string siteUrl, string userName, string password, string targetLibrary, MemoryStream file, string fileName, string envelopeID, bool createNewFolder)
        {
            using (ClientContext context = new ClientContext(siteUrl))
            {
                //Console.WriteLine("1. Connecting to SharePoint Online....");

                bool isOffice365 = false;
                Boolean.TryParse(ConfigurationManager.AppSettings["IsOffice365"], out isOffice365);
                if (isOffice365)
                {
                    SecureString securePassword = new SecureString();
                    foreach (char c in password.ToCharArray()) securePassword.AppendChar(c);

                    SharePointOnlineCredentials onlineCredentials = new SharePointOnlineCredentials(userName, securePassword);
                    context.Credentials = onlineCredentials;
                }
                else
                {
                    // OnPrem environment
                    return false;
                }

                Web web = context.Web;
                context.Load(web);
                context.ExecuteQuery();

                Microsoft.SharePoint.Client.List list = context.Web.Lists.GetByTitle(targetLibrary);
                context.Load(list.RootFolder);
                context.ExecuteQuery();

                // create a new folder named envelopeID
                if (createNewFolder)
                {
                    CreateFolder(web, list, list.RootFolder.ServerRelativeUrl, envelopeID);
                }


                //Console.WriteLine("Uploading envelope to SharePoint Online {0}....", siteUrl + @"/" + targetLibrary + @"/" + envelopeID + @"/");

                string fileUrl = string.Format("{0}/{1}", list.RootFolder.ServerRelativeUrl + @"/" + envelopeID + @"/", fileName);
                Microsoft.SharePoint.Client.File.SaveBinaryDirect(context, fileUrl, file, true);
            }
            return true;
        }
    }
}
