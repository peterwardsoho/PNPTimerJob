using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.SharePoint.Client;
using OfficeDevPnP.Core.Framework.TimerJobs;
using System.Configuration;
using OfficeDevPnP.Core;
using System.Security;
using Microsoft.Online.SharePoint.TenantAdministration;
using System.Xml;

namespace PNPTimer
{
    public class Job : TimerJob
    {
        public Job() : base("Job")
        {
            TimerJobRun += Job_TimerJobRun;
        }

        void Job_TimerJobRun(object sender, TimerJobRunEventArgs e)
        {
            Console.WriteLine("----- Timer Job Triggering ---");
            try
            {
                XmlReaderSettings settings = new XmlReaderSettings();
                settings.DtdProcessing = DtdProcessing.Parse;
                e.WebClientContext.Load(e.WebClientContext.Web);
                e.WebClientContext.ExecuteQueryRetry();
                Web web = e.WebClientContext.Web;
                Console.WriteLine(web.Title);
                e.WebClientContext.Load(web);
                e.WebClientContext.ExecuteQuery();
                Console.WriteLine("Site Request");
                List list = web.Lists.GetByTitle("Site Request");
                e.WebClientContext.Load(list);
                e.WebClientContext.ExecuteQuery();
                GetSiteCollectionDetails(e.WebClientContext, list);
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.ToString());
            }

        }

        private static void GetSiteCollectionDetails(ClientContext ctx, List list)
        {
            try
            {
                var camlQuery = new CamlQuery() { ViewXml = "<View><Query><Where><Eq><FieldRef Name='Status' />"+
                    "<Value Type='Choice'>Requested</Value></Eq></Where></Query></View>" };
                var listCollection = list.GetItems(camlQuery);
                ctx.Load(listCollection);
                ctx.ExecuteQuery();
                foreach (ListItem item in listCollection)
                {
                    string ID = item["ID"].ToString();
                    string Title = item["Title"].ToString();
                    string Template = item["Template"].ToString();
                    var fieldValues = (FieldUserValue)item.FieldValues["SiteOwner"];
                    var SiteOwner = ctx.Web.SiteUsers.GetById(fieldValues.LookupId);
                    ctx.Load(SiteOwner, x => x.Email);
                    ctx.ExecuteQuery();
                    string SiteOwnerEmail = SiteOwner.Email.ToString();
                    string Status = item["Status"].ToString();
                    if (Status == "Requested")
                    {
                        string url = ConfigurationManager.AppSettings["rootSiteUrl"] + ConfigurationManager.AppSettings["managedPath"] + Title;
                        if (!CheckSiteCollectionExist(ctx, url))
                        {
                            CreateSiteCollection(Title, Template, SiteOwnerEmail, url);
                            item["Status"] = "Approved";
                            item["Link"] = url;
                            item.Update();
                            ctx.ExecuteQuery();
                            Console.ReadKey();
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.ToString());
            }
        }

        private static bool CheckSiteCollectionExist(ClientContext ctx, string url)
        {
            if (ctx.WebExistsFullUrl(url))
            {
                Console.WriteLine(url + "Site Exist");
                return true;
            }
            else
            {
                return false;
            }
        }

        private static void CreateSiteCollection(string Title, string Template, string SiteOwnerEmail, string url)
        {
            try
            {
                string UserName = ConfigurationManager.AppSettings["UserName"];
                string Password = ConfigurationManager.AppSettings["Password"];
                string tenantAdminUrl = ConfigurationManager.AppSettings["tenantAdminUrl"];
                using (var clientContext = new ClientContext(tenantAdminUrl))
                {
                    SecureString passWord = new SecureString();
                    foreach (char c in Password.ToCharArray()) passWord.AppendChar(c);
                    clientContext.Credentials = new SharePointOnlineCredentials(UserName, passWord);
                    var tenant = new Tenant(clientContext);
                    var siteCreationProperties = new SiteCreationProperties();
                    siteCreationProperties.Url = url;
                    siteCreationProperties.Template = (Template == "Team Site(Modern)") ? "SITEPAGEPUBLISHING#0" : "";
                    siteCreationProperties.Title = Title;
                    siteCreationProperties.Owner = SiteOwnerEmail;
                    SpoOperation spo = tenant.CreateSite(siteCreationProperties);
                    clientContext.Load(tenant);
                    clientContext.Load(spo, i => i.IsComplete);
                    clientContext.ExecuteQuery();
                    while (!spo.IsComplete)
                    {
                        System.Threading.Thread.Sleep(30000);
                        spo.RefreshLoad();
                        clientContext.ExecuteQuery();
                    }
                    Console.WriteLine("SiteCollection Created.");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.ToString());
            }




        }

    }

    
}
