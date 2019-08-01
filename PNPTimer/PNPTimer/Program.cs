using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.SharePoint.Client;
using OfficeDevPnP.Core.Framework.TimerJobs;
using System.Security;
using OfficeDevPnP.Core.Sites;
using Microsoft.Online.SharePoint.TenantAdministration;
using OfficeDevPnP.Core;
using System.Configuration;
using PNPTimer;

namespace PNPTimer
{
    class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("Starting the Console");
            Job job = new Job();           
            var UserName = ConfigurationManager.AppSettings["UserName"];
            var Password = ConfigurationManager.AppSettings["Password"];
            Console.WriteLine(UserName);           
            job.UseOffice365Authentication(UserName, Password);
            job.AddSite(ConfigurationManager.AppSettings["targetSiteURL"]);
            job.Run();
        }
    }
}
