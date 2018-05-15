using System;
using System.Configuration;
using System.IO;
using DocuSign.eSign.Api;
using DocuSign.eSign.Model;
using DocuSign.eSign.Client;

namespace MSBuild2018
{
    class DocuSignEnvelopeHelper
    {
        /// <summary>
        /// UploadDocuSignEnvelopeToSharePoint - uploads a COMPLETED DocuSign envelope to a Sharepoint folder named as envelopeID
        /// </summary>
        /// <param name="accountID"></param>
        /// <param name="envelopeID"></param>
        static public void UploadDocuSignEnvelopeToSharePoint(string accountID, string envelopeID)
        {
            string userId = ConfigurationManager.AppSettings["UserId"];
            string oauthBasePath = ConfigurationManager.AppSettings["OAuthBasePath"];
            string integratorKey = ConfigurationManager.AppSettings["IntegratorKey"];
            string privateKeyFilename = AppContext.BaseDirectory + "PrivateKey.txt";
            string host = ConfigurationManager.AppSettings["Host"];
            int expiresInHours = 1;

            string siteUrl = ConfigurationManager.AppSettings["SiteUrl"];
            string targetLibrary = ConfigurationManager.AppSettings["TargetLibrary"];
            string userName = ConfigurationManager.AppSettings["UserName"];
            string password = ConfigurationManager.AppSettings["Password"];

            ApiClient apiClient = new ApiClient(host);
            apiClient.ConfigureJwtAuthorizationFlow(integratorKey, userId, oauthBasePath, privateKeyFilename, expiresInHours);

            /////////////////////////////////////////////////////////////////
            // STEP 1: LOGIN API        
            /////////////////////////////////////////////////////////////////
            AuthenticationApi authApi = new AuthenticationApi(apiClient.Configuration);
            LoginInformation loginInfo = authApi.Login();

            // find the default account for this user
            foreach (LoginAccount loginAcct in loginInfo.LoginAccounts)
            {
                if (loginAcct.IsDefault == "true")
                {
                    accountID = loginAcct.AccountId;

                    string[] separatingStrings = { "/v2" };

                    // Update ApiClient with the new base url from login call
                    apiClient = new ApiClient(loginAcct.BaseUrl.Split(separatingStrings, StringSplitOptions.RemoveEmptyEntries)[0]);
                    break;
                }
            }

            /////////////////////////////////////////////////////////////////
            // STEP 2: GET DOCUMENTS API        
            /////////////////////////////////////////////////////////////////

            EnvelopesApi envelopesApi = new EnvelopesApi(apiClient.Configuration);
            Envelope envInfo = envelopesApi.GetEnvelope(accountID, envelopeID);

            if (envInfo.Status.ToLower().CompareTo("completed") == 0)
            {
                // upload all documents for accountID and envelopeID
                MemoryStream docStream = GetAllDocuments(accountID, envelopeID);
                Console.WriteLine("Uploading to SharePoint all documents for envelope {0}", envelopeID);
                string fileName = envInfo.EnvelopeId + ".pdf";
                SharePointOnlineHelper.UploadFileStreamToSharePoint(siteUrl, userName, password, targetLibrary, docStream, fileName, envelopeID, false);

                // upload Certificate of Completion for accountID and envelopeID
                Console.WriteLine("Uploading to SharePoint the Certificate Of Completion for envelope {0}", envelopeID);
                MemoryStream cocStream = GetCertificateOfCompletion(accountID, envelopeID);
                fileName = "COC_" + envInfo.EnvelopeId + ".pdf";
                SharePointOnlineHelper.UploadFileStreamToSharePoint(siteUrl, userName, password, targetLibrary, cocStream, fileName, envelopeID, false);
            }
            else
            {
                Console.WriteLine("Download DocuSign documents can be performed only for COMPLETED envelopes.");
            }
        }

        /// <summary>
        /// GetDocumentByName: receives an envelope ID and a document name and returns the document as MemoryStream from the corresponding envelope.
        /// </summary>
        /// <param name="accountID">Account ID, as string</param>
        /// <param name="envelopeID">Envelope ID, as string</param>
        /// <param name="docName">Document name, as string</param>
        /// <returns>
        /// MemoryStream as document with docName in envelope with EnvelopeID
        /// </returns>
        public static MemoryStream GetDocumentByName(string accountID,
                                                    string envelopeID,
                                                    string docName)
        {
            MemoryStream docStream = null;

            // |EnvelopesApi| contains methods related to envelopes and envelope recipients
            EnvelopesApi envelopesApi = new EnvelopesApi();
            EnvelopeDocumentsResult docsList = envelopesApi.ListDocuments(accountID, envelopeID);

            // read how many documents are in the envelope
            int docCount = docsList.EnvelopeDocuments.Count;

            // loop through the envelope's documents, search for document with DocName and download it.
            for (int i = 0; i < docCount; i++)
            {
                if (docsList.EnvelopeDocuments[i].Name.Equals(docName))
                {
                    // GetDocument() API call returns a MemoryStream
                    docStream = (MemoryStream)envelopesApi.GetDocument(accountID, envelopeID, docsList.EnvelopeDocuments[i].DocumentId);
                    break;
                }
            }

            return docStream;
        }

        /// <summary>
        /// GetDocumentByID: the function receives an envelope ID and a document ID and returns the MemoryStream of a document according to document ID.
        /// The certificate is not included in the FileStream object.
        /// </summary>
        /// <param name="accountID">Account ID, as string</param>
        /// <param name="envelopeID">Envelope ID, as string</param>
        /// <param name="documentID">Document ID, as string</param>
        /// <returns>
        /// MemoryStream as all documents in envelope with envelopeID filtered by documentID combined
        /// </returns>
        public static MemoryStream GetDocumentByID(string accountID,
                                                    string envelopeID,
                                                    string documentID)
        {
            MemoryStream docStream = null;

            // |EnvelopesApi| contains methods related to envelopes and envelope recipients
            EnvelopesApi envelopesApi = new EnvelopesApi();

            docStream = (MemoryStream)envelopesApi.GetDocument(accountID, envelopeID, documentID);

            return docStream;
        }

        /// <summary>
        /// GetCertificateOfCompletion: receives an envelope ID and account ID and returns the Certificate of Completion for the envelopeID.
        /// </summary>
        /// <param name="accountID">Account ID, as string</param>
        /// <param name="envelopeID">Envelope ID, as string</param>
        /// <returns>
        /// FileStream for the Certificate Of Compeltion for envelopeID
        /// </returns>
        public static MemoryStream GetCertificateOfCompletion( string accountID,
                                                                string envelopeID)
        {
            MemoryStream docStream = null;

            // |EnvelopesApi| contains methods related to envelopes and envelope recipients
            EnvelopesApi envelopesApi = new EnvelopesApi();

            docStream = (MemoryStream)envelopesApi.GetDocument(accountID, envelopeID, "certificate");
            return docStream;
        }

        /// <summary>
        /// GetAllDocuments: receives an envelope ID and account ID and returns the combined MemoryStream of all documents according to envelopeID.
        /// According to includeCoC parameter, the Certificate of Completion will  be or will not be included in the resulting MemoryStream.
        /// </summary>
        /// <param name="accountID">Account ID, as string</param>
        /// <param name="envelopeID">Envelope ID, as string</param>
        /// <param name="includeCoC">Resulting stream includes the CoC, as boolean</param>
        /// <returns>
        /// MemoryStream as all documents in envelope with envelopeID
        /// </returns>
        public static MemoryStream GetAllDocuments(string accountID,
                                                    string envelopeID)
        {
            MemoryStream docStream = null;

            // |EnvelopesApi| contains methods related to envelopes and envelope recipients
            EnvelopesApi envelopesApi = new EnvelopesApi();

            docStream = (MemoryStream)envelopesApi.GetDocument(accountID, envelopeID, "combined");

            return docStream;
        }
    }
}
