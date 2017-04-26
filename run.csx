#r "Microsoft.SharePoint.Client.dll"  
#r "Microsoft.SharePoint.Client.Runtime.dll"  
#r "OfficeDevPnP.Core.dll"  
#r "System.XML.dll"
#r "System.Xml.Linq.dll"
#r "Microsoft.SharePoint.Client.Taxonomy.dll"
#r "Microsoft.SharePoint.Client.Publishing.dll"

using System.Net;
using Microsoft.SharePoint.Client;
using OfficeDevPnP.Core.Framework.Provisioning.Connectors;
using OfficeDevPnP.Core.Framework.Provisioning.Model;
using OfficeDevPnP.Core.Framework.Provisioning.ObjectHandlers;
using OfficeDevPnP.Core.Framework.Provisioning.Providers.Xml;
using System.Xml;
using System.Xml.Linq;
using System.Linq;
using System.Security;


public static async Task<HttpResponseMessage> Run(HttpRequestMessage req, TraceWriter log)
{
    log.Info("C# HTTP trigger function processed a request.");

    string siteUrl = "https://smbconcurrency.sharepoint.com/sites/globalnav/";  
    string userName = "admin@smbconcurrency.onmicrosoft.com";  
    string password = "Breanna99";  
    OfficeDevPnP.Core.AuthenticationManager authManager = new OfficeDevPnP.Core.AuthenticationManager();  
    try  
    {  
        // Get and set the client context  
        // Connects to SharePoint online site using inputs provided  
        using (var clientContext = authManager.GetSharePointOnlineAuthenticatedContextTenant(siteUrl, userName, password))  
        {  
  
            // List Name input  
            string listName = "GlobalNavigationSites";  
            // Retrieves list object using title  
            List list = clientContext.Site.RootWeb.GetListByTitle(listName);  

            string templateWebUrl = "https://smbconcurrency.sharepoint.com/sites/globalnav";
            string targetWebUrl ="https://smbconcurrency.sharepoint.com/sites/globalnav";
            
            string pwdS = "Breanna99";
            SecureString pwd = new SecureString();
            foreach (char c in pwdS.ToCharArray()) pwd.AppendChar(c);

            // GET the template from existing site and serialize
            // Serializing the template for later reuse is optional
            ProvisioningTemplate template = GetProvisioningTemplate(templateWebUrl, userName, pwd, log);

            if (list != null)  
            {  
                // Apply Provisioning to site collections in this list
                // TOOD - compare all site collections against ones in this list and then process the
                // ones that have ApplyNav enabled
                ApplyProvisioningTemplate(targetWebUrl, userName, pwd, log);

                // Returns required result  
                return req.CreateResponse(HttpStatusCode.OK, list.Id);              
            }  
            else  
            {  
                log.Info("List is not available on the site");  
            }  
  
        }  
    }  
    catch (Exception ex)  
    {  
        log.Info("Error Message: " + ex.Message);  
    }  

    return req.CreateResponse(HttpStatusCode.OK, "End of function");  
      
}   
 private static ProvisioningTemplate GetProvisioningTemplate(string webUrl, string userName, System.Security.SecureString pwd, TraceWriter log)
        {
            using (var ctx = new ClientContext(webUrl))
            {
                ctx.Credentials = new SharePointOnlineCredentials(userName, pwd);
                ctx.RequestTimeout = System.Threading.Timeout.Infinite;

                // Just to output the site details
                Web web = ctx.Web;
                ctx.Load(web, w => w.Title);
                ctx.ExecuteQueryRetry();

                //Console.ForegroundColor = ConsoleColor.White;
               log.Info("Your site title is: " + ctx.Web.Title);
               // Console.ForegroundColor = defaultForeground;

                ProvisioningTemplateCreationInformation ptci
                        = new ProvisioningTemplateCreationInformation(ctx.Web);

                // Create FileSystemConnector to store a temporary copy of the template 
                ptci.FileConnector = new FileSystemConnector(@"D:\home\site\wwwroot\HttpTriggerATCNav\GlobalNav", "");
                ptci.PersistBrandingFiles = true;
                ptci.HandlersToProcess = Handlers.Navigation;

                ptci.ProgressDelegate = delegate (String message, Int32 progress, Int32 total)
                {
                    // Only to output progress for console UI
                    log.Info(string.Format("{0:00}/{1:00} - {2}", progress, total, message));
                };

                // Execute actual extraction of the template
                ProvisioningTemplate template = ctx.Web.GetProvisioningTemplate(ptci);


                // We can serialize this template to save and reuse it
                // Optional step 
                XMLTemplateProvider provider =
                        new XMLFileSystemTemplateProvider(@"D:\home\site\wwwroot\HttpTriggerATCNav\GlobalNav", "");
                provider.SaveAs(template, "PnPProvisioningDemo.xml");

                log.Info("The existing site template has been extracted and saved.");

                // Get Navigation only in Site Provisioning Template
                try
                {                   

                    // Load site provisioning template (Navigation only)
                    XDocument doc = XDocument.Load(@"D:\home\site\wwwroot\HttpTriggerATCNav\GlobalNav\PnPProvisioningDemo.xml");

                    // Get Current Navigation Nodes
                    var currNavNodes = from node in doc.Descendants(doc.Root.GetNamespaceOfPrefix("pnp") + "CurrentNavigation")
                                       select node;

                    // Remove Current Navigation
                    currNavNodes.Remove();

                    // Save new provisioning file with only Global Navigation in the Navigation Nodes
                 //   doc.Save(@"D:\home\site\wwwroot\HttpTriggerATCNav\GlobalNav\GlobalNav.xml");

                    log.Info("The site template has been processed to only contain Global Navigation nodes and saved.");
                
                }
                catch (Exception ex)
                {
                    //Console.ForegroundColor = ConsoleColor.White;
                   log.Info("There were no Navigation nodes found in the template for site: " + ctx.Web.Title);
                   log.Info("Error: " + ex.InnerException.ToString());
                   // Console.ForegroundColor = defaultForeground;
                }


                return template;
            }
        }


         private static void ApplyProvisioningTemplate(string webUrl, string userName, System.Security.SecureString pwd, TraceWriter log)
        {
            using (var ctx = new ClientContext(webUrl))
            {
                // ctx.Credentials = new NetworkCredentials(userName, pwd);
                ctx.Credentials = new SharePointOnlineCredentials(userName, pwd);
                ctx.RequestTimeout = System.Threading.Timeout.Infinite;

                // Just to output the site details
                Web web = ctx.Web;
                ctx.Load(web, w => w.Title);
                ctx.ExecuteQueryRetry();

                // Configure the XML file system provider
                XMLTemplateProvider providerNewNav =
                new XMLFileSystemTemplateProvider(@"D:\home\site\wwwroot\HttpTriggerATCNav\GlobalNav", "");

                // Load the template from the XML stored copy
                ProvisioningTemplate templateNewNav = providerNewNav.GetTemplate("GlobalNav.xml");

                // start timer
                log.Info(string.Format("Start Applying Template: {0:hh.mm.ss}", DateTime.Now));

                // Apply the template to another site
                var applyingInformation = new ProvisioningTemplateApplyingInformation();

                // overwrite and remove existing navigation nodes
                applyingInformation.ClearNavigation = true;

                applyingInformation.ProgressDelegate = (message, step, total) =>
                {
                   log.Info(string.Format("{0}/{1} Provisioning {2}", step, total, message));
                };

                // Apply the template to the site
                web.ApplyProvisioningTemplate(templateNewNav, applyingInformation);

                log.Info(string.Format("Done applying template: {0:hh.mm.ss}", DateTime.Now));
            }
        }