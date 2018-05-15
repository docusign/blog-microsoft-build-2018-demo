using System;
using System.Collections.Generic;
using System.Configuration;
using DocuSign.eSign.Api;
using DocuSign.eSign.Model;
using DocuSign.eSign.Client;


namespace MSBuild2018
{
    class DocuSignDemo
    {
        static string accountID = string.Empty;
        static void Main(string[] args)
        {
            #region DocuSign Envelope
            // create DocuSign envelope
            Console.WriteLine("Creating a DocuSign envelope...");
            string envelopeID = CreateDocuSignEnvelope();
            Console.WriteLine("Envelope ID sent to " + ConfigurationManager.AppSettings["RecipientEmail"] + ": " + envelopeID);
            if ((bool)Convert.ToBoolean(ConfigurationManager.AppSettings["ShowSharePointDemo"]) == true)
                //Configured to run SharePoint demo - prompt accordingly
                Console.WriteLine("After envelope is completed (signed), press any key to continue to upload to SharePoint...");
            else
                //Not configured to run SharePoint demo
                Console.WriteLine("Press any key to exit...");

            Console.ReadKey();
            #endregion

            #region SharePoint
            if ((bool)Convert.ToBoolean(ConfigurationManager.AppSettings["ShowSharePointDemo"]) == true)
            {
                Console.WriteLine("Uploading COMPLETED envelope {0} to SharePoint...", envelopeID);
                // create completed envelope external repository folder in SharePoint
                SharePointOnlineHelper.CreateFolderForEnvelope(envelopeID);

                DocuSignEnvelopeHelper.UploadDocuSignEnvelopeToSharePoint(accountID, envelopeID);
                Console.WriteLine("Done uploading envelope {0} to SharePoint", envelopeID);
                Console.WriteLine("Press any key to exit...");
                Console.ReadKey();
            }
            #endregion
        }

        static string CreateDocuSignEnvelope()
        {
            string userId = ConfigurationManager.AppSettings["UserId"];
            string oauthBasePath = ConfigurationManager.AppSettings["OAuthBasePath"];
            string integratorKey = ConfigurationManager.AppSettings["IntegratorKey"];
            string privateKeyFilename = AppContext.BaseDirectory + "PrivateKey.txt";
            string host = ConfigurationManager.AppSettings["Host"];
            string templateId = ConfigurationManager.AppSettings["TemplateID"];
            int expiresInHours = 1;

            //string accountId = string.Empty;

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
            // STEP 2: CREATE ENVELOPE API        
            /////////////////////////////////////////////////////////////////				

            EnvelopeDefinition envDef = new EnvelopeDefinition();
            envDef.EmailSubject = "MS Build Demo - Please sign this doc";

            // assign recipient to template role by setting name, email, and role name.  Note that the
            // template role name must match the placeholder role name saved in your account template.  
            TemplateRole tRole = new TemplateRole();
            tRole.Email = ConfigurationManager.AppSettings["RecipientEmail"];
            tRole.Name = ConfigurationManager.AppSettings["RecipientName"];
            tRole.RoleName = "Signer";
            List<TemplateRole> rolesList = new List<TemplateRole>() { tRole };

            // add the role to the envelope and assign valid templateId from your account
            envDef.TemplateRoles = rolesList;
            envDef.TemplateId = templateId;

            // set envelope status to "sent" to immediately send the signature request
            envDef.Status = "sent";

            // |EnvelopesApi| contains methods related to creating and sending Envelopes (aka signature requests)
            EnvelopesApi envelopesApi = new EnvelopesApi(apiClient.Configuration);
            EnvelopeSummary envelopeSummary = envelopesApi.CreateEnvelope(accountID, envDef);

            return envelopeSummary.EnvelopeId;
        }
    }
}
