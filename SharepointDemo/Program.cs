using Microsoft.SharePoint.Client;
using Microsoft.SharePoint.Client.WorkflowServices;
using System;
using System.Net;
using System.Security;

namespace SharepointDemo
{
    class Program
    {
        static void Main(string[] args)
        {
            var siteUrl = "https://O360DEV01-SP.chrlab.com/sites/ABC-P";
            var domain = "";
            var username = "";
            var password = "";

            var context = new ClientContext(siteUrl);
            context.Credentials = new NetworkCredential(username, GetSecureString(password), domain); //new SharePointOnlineCredentials(username, GetSecureString(password));

            var web = context.Web;

            context.Load(web);
            context.ExecuteQuery();
            Console.WriteLine($"Web Title: {web.Title}");

            GetAllLists(context);

            GetWorkflowInstances(context);
            GetWorkflowTasks(context);
            GetOmnia360Orders(context);

            //context.Load(web, w => w.Title, w => w.Description);
            //context.ExecuteQuery();
            //Console.WriteLine($"Web Title:   {web.Title}");
            //Console.WriteLine($"Description: {web.Description}");


            //// connect with rest api
            //var siteUrl = "sp.chrsolutions.com";
            //var baseSharepointUri = new Uri($"https://{siteUrl}/_api_site");
            //using var httpClient = new HttpClient();

            Console.WriteLine("\nPress ENTER to quit");
            Console.ReadLine();
        }

        static void GetWorkflowInstances(ClientContext context)
        {
            var workflowServicesManager = new WorkflowServicesManager(context, context.Web);
            var subscriptionService = workflowServicesManager.GetWorkflowSubscriptionService();
            var instanceService = workflowServicesManager.GetWorkflowInstanceService();

            //??? what is this -> subscriptionService.RegisterInterestInList(listId, eventName)

            //ActivityExecutionStatus x;
            //System.Activities.Activity x;

            // todo: look at what it takes to use the Windows Workflow namespaces from .NET on Sharepoint workflows
            // System.Activities, ...

            Console.WriteLine("\n=================================================================================");
            Console.WriteLine("Workflow Instances");
            Console.WriteLine("=================================================================================");
            var subscriptions = subscriptionService.EnumerateSubscriptions();
            context.Load(subscriptions);
            context.ExecuteQuery();
            foreach (var subscription in subscriptions)
            {
                context.Load(subscription);
                context.ExecuteQuery();

                var clientResult = instanceService.CountInstances(subscription);
                context.ExecuteQuery();

                Console.WriteLine($"SubscriptionName:{subscription.Name} InstanceCount:{clientResult.Value} StatusFldName:{subscription.StatusFieldName}");
            }
        }

        static void GetWorkflowTasks(ClientContext context)
        {
            var workflowTaskList = context.Web.Lists.GetByTitle("Workflow Tasks");
            var camlQuery = new CamlQuery
            {
                ViewXml = "<View><RowLimit>100</RowLimit></View>"
            };
            var collListItem = workflowTaskList.GetItems(camlQuery);
            context.Load(collListItem,
                items => items.Include(
                    item => item.Id,
                    item => item.DisplayName,
                    item => item.HasUniqueRoleAssignments));

            context.ExecuteQuery();


            Console.WriteLine("\n=================================================================================");
            Console.WriteLine("Workflow Tasks");
            Console.WriteLine("=================================================================================");
            foreach (var item in collListItem)
            {
                Console.WriteLine($"ID:{item.Id}  Display Name:{item.DisplayName}  Unique Roles:{item.HasUniqueRoleAssignments}");
            }
        }

        static void GetOmnia360Orders(ClientContext context)
        {
            var workflowTaskList = context.Web.Lists.GetByTitle("Omnia360 Orders");
            var camlQuery = new CamlQuery
            {
                ViewXml = "<View><RowLimit>100</RowLimit></View>"
            };
            var collListItem = workflowTaskList.GetItems(camlQuery);
            context.Load(collListItem
                //,
                //items => items.Include(
                //    item => item.Id,
                //    item => item.DisplayName,
                //    item => item.HasUniqueRoleAssignments)
            );

            context.ExecuteQuery();


            Console.WriteLine("\n=================================================================================");
            Console.WriteLine("Omnia360 Orders");
            Console.WriteLine("=================================================================================");
            foreach (var item in collListItem)
            {
                Console.WriteLine($"ID:{item.Id}  OrderNumber:{item.FieldValues["OrderNumber"]}  WkflowInstanceId:{item.FieldValues["WorkflowInstanceID"]}  WkflowType:{item.FieldValues["WorkflowType"]}");
            }

            // NOTE: Field Values conatain the following keys
            // [0]: "ContentTypeId"
            // [1]: "Title"
            // [2]: "_ModerationComments"
            // [3]: "File_x0020_Type"
            // [4]: "OrderNumber"
            // [5]: "OrderId"
            // [6]: "WorkflowType"
            // [7]: "IntegrationId"
            // [8]: "OrganizationName"
            // [9]: "Omnia360ApiUrl"
            // [10]: "ProvisioningDate"
            // [11]: "Billing_x0020_Only"
            // [12]: "Facility_x0020_Changed_x0020_Com"
            // [13]: "SO_x0020_Training_x0020_Complete"
            // [14]: "Billing_x0020_Only_x0020_1"
            // [15]: "Eleeson_x0020_Test"
            // [16]: "Conditional_x0020_workflow_x0020"
            // [17]: "CP_x0020_Complete_x0020_Workflow"
            // [18]: "Provision_x0020_and_x0020_Activa"
            // [19]: "ID"
            // [20]: "Modified"
            // [21]: "Created"
            // [22]: "Author"
            // [23]: "Editor"
            // [24]: "_HasCopyDestinations"
            // [25]: "_CopySource"
            // [26]: "owshiddenversion"
            // [27]: "WorkflowVersion"
            // [28]: "_UIVersion"
            // [29]: "_UIVersionString"
            // [30]: "Attachments"
            // [31]: "_ModerationStatus"
            // [32]: "InstanceID"
            // [33]: "Order"
            // [34]: "GUID"
            // [35]: "WorkflowInstanceID"
            // [36]: "FileRef"
            // [37]: "FileDirRef"
            // [38]: "Last_x0020_Modified"
            // [39]: "Created_x0020_Date"
            // [40]: "FSObjType"
            // [41]: "SortBehavior"
            // [42]: "FileLeafRef"
            // [43]: "UniqueId"
            // [44]: "SyncClientId"
            // [45]: "ProgId"
            // [46]: "ScopeId"
            // [47]: "MetaInfo"
            // [48]: "_Level"
            // [49]: "_IsCurrentVersion"
            // [50]: "ItemChildCount"
            // [51]: "FolderChildCount"
            // [52]: "AppAuthor"
            // [53]: "AppEditor"
        }

        static void GetAllLists(ClientContext context)
        {
            var web = context.Web;

            context.Load(web.Lists,
                lists => lists.Include(list => list.Title,
                         list => list.Id));

            context.ExecuteQuery();

            foreach (var list in web.Lists)
            {
                Console.WriteLine($"Title: {list.Title}");
            }
        }

        static SecureString GetSecureString(string unsecureString)
        {
            var secureString = new SecureString();

            foreach (var c in unsecureString.ToCharArray())
            {
                secureString.AppendChar(c);
            }

            return secureString;
        }
    }
}
