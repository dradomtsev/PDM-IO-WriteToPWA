using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using Microsoft.SharePoint.Client;
using Microsoft.ProjectServer.Client;

using System.Net;

namespace PDM.IO.PWA.Assignments
{
    class Program
    {
        private const string SiteURL = "https://archimatika.sharepoint.com/sites/pwatest";
        private static ProjectContext projContext = new ProjectContext(SiteURL);
        static void Main()
        {
            WorkPWA();
        }
        static void WorkPWA()
        {
            // Connect to Sharepoint using cookies
            var cookies = PDM.IO.PWA.Login.WebLogin.GetWebLoginCookie(new Uri(SiteURL));
            projContext.ExecutingWebRequest += delegate (object sender, WebRequestEventArgs e)
            {
                e.WebRequestExecutor.WebRequest.CookieContainer = new CookieContainer();
                e.WebRequestExecutor.WebRequest.CookieContainer.SetCookies(new Uri(SiteURL), cookies);
            };

            // Write data to PWA
            SetResourceDefaultAssignmentOwner();
        }

        private static void SetResourceDefaultAssignmentOwner()
        {
            //bool bParse = Enum.TryParse(stResourceType, out ertType);
            EnterpriseResourceCollection entresources = projContext.EnterpriseResources;
            projContext.Load(entresources, item => item.Where<EnterpriseResource>(i => /*i.ResourceType == EnterpriseResourceType.Work && */i.Name == "TestResource-001").Include(p => p.Id, p => p.Name, p => p.User, p => p.DefaultAssignmentOwner));
            try
            {
                projContext.ExecuteQuery();
                //projContext.ExecuteQueryAsync().Wait();
            }
            catch (System.NullReferenceException e)
            {
                Console.WriteLine(e.Message);
            }
            EnterpriseResource resource = entresources.FirstOrDefault();


            EnterpriseResourceCollection entresourcesDefAssignmentOwner = projContext.EnterpriseResources;
            projContext.Load(entresourcesDefAssignmentOwner, item => item.Where<EnterpriseResource>(i => /*i.ResourceType == EnterpriseResourceType.Work && */i.Name == "Olha Hannochka").Include(p => p.Id, p => p.User));// && i.Name == "Olha Hannochka" ) );
            projContext.ExecuteQuery();

            EnterpriseResource resourceDefAssignmentOwner = entresourcesDefAssignmentOwner.FirstOrDefault();

            resource.DefaultAssignmentOwner = resourceDefAssignmentOwner.User;
            
            projContext.EnterpriseResources.Update();
            projContext.ExecuteQuery();
        }
    }
}
