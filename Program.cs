using AccessTeamsReports.Models;
using AccessTeamsReports.Utilities;
using Microsoft.Graph;
using Microsoft.Graph.Auth;
using Microsoft.Identity.Client;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AccessTeamsReports
{
    class Program
    {
        static async Task Main(string[] args)
        {
            IPublicClientApplication publicClientApplication = PublicClientApplicationBuilder
                        .Create("debdb95b-0c44-4e47-a97e-5a4d7b9291f9")
                        .WithAuthority("https://login.microsoftonline.com/d7037d74-e72d-44e9-8f94-c0b2b4845174")
                        .WithDefaultRedirectUri()
                        .Build();
           
            string[] scopes = new string[] { "Organization.ReadWrite.All","Reports.Read.All" };
            var accounts = await publicClientApplication.GetAccountsAsync();
            
            AuthenticationResult result;
            try
            {
                result = await publicClientApplication.AcquireTokenSilent(scopes, accounts.FirstOrDefault())
                            .ExecuteAsync();
            }
            catch (MsalUiRequiredException)
            {
                result = await publicClientApplication.AcquireTokenInteractive(scopes)
                            .ExecuteAsync();
            }

            DeviceCodeProvider authProvider = new DeviceCodeProvider(publicClientApplication, scopes);

            // TODO: MS Not implemented this feature on the Reports APIs (Only CSV at this point)
            //var queryOptions = new List<QueryOption>()
            //    {
            //        new QueryOption("$format", "application/json")
            //    };

            // Adding a custom serializer to handle CSV and Transmogrify into JSON (check ReportCustomSerializers for specific Report Serializer)
            HttpProvider provider = new HttpProvider(ReportCustomSerializersOptions.teamsUserActivityUserCounts);
            GraphServiceClient graphClient = new GraphServiceClient(authProvider, provider);

            var x = await graphClient.Reports.GetTeamsUserActivityCounts("D30").Request().GetAsync();
                       
            using (StreamReader sw = new StreamReader(x.Content))
            {
                string json = sw.ReadToEnd();
                Console.Write(json);
            }

            Console.Read();


        }
    }
}
