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
        private static ProjectContext objContext = new ProjectContext(SiteURL);
        const int PROJECT_BLOCK_SIZE = 20;
        static void Main()
        {
            WorkPWA();
        }
        static void WorkPWA()
        {
            // Connect to Sharepoint using cookies
            var cookies = PDM.IO.PWA.Login.WebLogin.GetWebLoginCookie(new Uri(SiteURL));
            objContext.ExecutingWebRequest += delegate (object sender, WebRequestEventArgs e)
            {
                e.WebRequestExecutor.WebRequest.CookieContainer = new CookieContainer();
                e.WebRequestExecutor.WebRequest.CookieContainer.SetCookies(new Uri(SiteURL), cookies);
            };

            // Write data to PWA
            //SetResourceDefaultAssignmentOwner();
            SetResourceAssignmentOwner();
        }

        private static void SetResourceDefaultAssignmentOwner()
        {
            EnterpriseResourceCollection entresources = objContext.EnterpriseResources;
            objContext.Load(entresources, item => item.Where<EnterpriseResource>(i => /*i.ResourceType == EnterpriseResourceType.Work && */i.Name == "TestResource-001").Include(p => p.Id, p => p.Name, p => p.User, p => p.DefaultAssignmentOwner));
            objContext.ExecuteQuery();
            EnterpriseResource resource = entresources.FirstOrDefault();


            EnterpriseResourceCollection entresourcesDefAssignmentOwner = objContext.EnterpriseResources;
            objContext.Load(entresourcesDefAssignmentOwner, item => item.Where<EnterpriseResource>(i => /*i.ResourceType == EnterpriseResourceType.Work && */i.Name == "Olha Hannochka").Include(p => p.Id, p => p.User));// && i.Name == "Olha Hannochka" ) );
            objContext.ExecuteQuery();
            EnterpriseResource resourceDefAssignmentOwner = entresourcesDefAssignmentOwner.FirstOrDefault();

            resource.DefaultAssignmentOwner = resourceDefAssignmentOwner.User;
            
            objContext.EnterpriseResources.Update();
            objContext.ExecuteQuery();
        }

        private static void SetResourceAssignmentOwner()
        {
            EnterpriseResourceCollection entresources = objContext.EnterpriseResources;
            objContext.Load(entresources, item => item.Where<EnterpriseResource>(i => /*i.ResourceType == EnterpriseResourceType.Work && */i.Name == "Dmytro Radomtsev").Include(p => p.Id, p => p.Name, p => p.User, p => p.DefaultAssignmentOwner));
            objContext.ExecuteQuery();
            EnterpriseResource resource = entresources.FirstOrDefault();


            EnterpriseResourceCollection entresourcesDefAssignmentOwner = objContext.EnterpriseResources;
            objContext.Load(entresourcesDefAssignmentOwner, item => item.Where<EnterpriseResource>(i => /*i.ResourceType == EnterpriseResourceType.Work && */i.Name == "Olha Hannochka").Include(p => p.Id, p => p.User));// && i.Name == "Olha Hannochka" ) );
            objContext.ExecuteQuery();
            EnterpriseResource resourceAssignmentOwner = entresourcesDefAssignmentOwner.FirstOrDefault();

            IEnumerable<PublishedProject> projects;

            objContext.Load(objContext.Projects, qp => qp.Include(qr => qr.Id));
            objContext.ExecuteQuery();

            try
            {
                Guid[] allIds = objContext.Projects.Select(p => p.Id).ToArray();
                int numBlocks = allIds.Length / PROJECT_BLOCK_SIZE + 1;
                for (int i = 0; i < numBlocks; i++)
                {
                    IEnumerable<Guid> idBlock = allIds.Skip(i * PROJECT_BLOCK_SIZE).Take(PROJECT_BLOCK_SIZE);
                    Guid[] block = new Guid[PROJECT_BLOCK_SIZE];
                    Array.Copy(idBlock.ToArray(), block, idBlock.Count());

                    projects = objContext.LoadQuery(
                        objContext.Projects
                        .Where(p =>   
                            p.Id == block[0] || p.Id == block[1] ||
                            p.Id == block[2] || p.Id == block[3] ||
                            p.Id == block[4] || p.Id == block[5] ||
                            p.Id == block[6] || p.Id == block[7] ||
                            p.Id == block[8] || p.Id == block[9] ||
                            p.Id == block[10] || p.Id == block[11] ||
                            p.Id == block[12] || p.Id == block[13] ||
                            p.Id == block[14] || p.Id == block[15] ||
                            p.Id == block[16] || p.Id == block[17] ||
                            p.Id == block[18] || p.Id == block[19]
                        )
                        .Include(p => p.Id,
                            p => p.Name,
                            p => p.Assignments
                        )
                    );

                    objContext.ExecuteQuery();
                }
            }
            catch (Microsoft.SharePoint.Client.ServerException ex)
            {
                Console.WriteLine(ex.Message);
            }
            catch (Microsoft.SharePoint.Client.CollectionNotInitializedException ex)
            {
                Console.WriteLine(ex.Message);
            }

            foreach (PublishedProject pubProj in objContext.Projects)
            {
                var draftProject = pubProj.CheckOut();
                IEnumerable<DraftAssignment> assignments = objContext.LoadQuery(draftProject.Assignments.Where(a => a.Resource.Id == resource.Id).Include(a => a.Id, a => a.Resource, a => a.Owner));
                objContext.ExecuteQuery();
                foreach (DraftAssignment assignment in assignments)
                    assignment.Owner = resourceAssignmentOwner.User;

                draftProject.Publish(true);
                objContext.ExecuteQuery();
            }
        }
    }
}
