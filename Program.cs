using AccessTeamsReports.Models;
using AccessTeamsReports.Utilities;
using Microsoft.Graph;
using Microsoft.Graph.Auth;
using Microsoft.Identity.Client;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AccessTeamsReports
{
    /// <summary>
    /// Written override default behaviour of Reporting APIs for Team Activity to Serialize to JSON
    /// Author: Mark Franco - Microsft Technology Centre (Toronto)
    /// Note: Sample code provided as-is, not for direct use into production (To be used as a learning tool)
    /// </summary>
    class Program
    {
        static async Task Main(string[] args)
        {

            #region Authentication work Here
            IPublicClientApplication publicClientApplication = PublicClientApplicationBuilder
                        .Create(ConfigurationManager.AppSettings["APP_CLIENT_ID"])
                        .WithAuthority(ConfigurationManager.AppSettings["AAD_AUTHORITY"])
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

            // Create the Autentication Provider with Proper Scopes for TEams Reporting APIs
            DeviceCodeProvider authProvider = new DeviceCodeProvider(publicClientApplication, scopes);

            #endregion

            #region Calling Reports APIs
            // Call the APIs with the newly formed Authentication Provider
            bool success = await PullTeamsReports(authProvider);

            if (success)
                Console.WriteLine("Jobs Finished!");
            else
                Console.WriteLine("Error!");

            #endregion

            Console.WriteLine("Press <ENTER> to close Application");
            Console.Read();


        }

        private async static Task<bool> PullTeamsReports(DeviceCodeProvider authProvider)
        {
            const string period = "D7"; // 7 Days Worth of Data. Other options (D7,D30,D90,D180)

            #region GetTeamsUserActivityUserCounts
            try
            {
                // Adding a custom serializer to handle CSV and Transmogrify into JSON (NOTE: check ReportCustomSerializers for specific Report Serializer)
                HttpProvider httpProvider = new HttpProvider(ReportCustomSerializersOptions.teamsUserActivityUserCountsSerializer);
                GraphServiceClient graphClient = new GraphServiceClient(authProvider, httpProvider);

                // TODO: MS Not implemented this feature on the Reports APIs (Only CSV at this point)
                //var queryOptions = new List<QueryOption>()
                //    {
                //        new QueryOption("$format", "application/json")
                //    };

                var response = await graphClient.Reports.GetTeamsUserActivityUserCounts(period).Request(/* queryOptions */).GetAsync();
                using (StreamReader sw = new StreamReader(response.Content))
                {
                    using (var file = new StreamWriter("c:\\temp\\GetTeamsUserActivityUserCounts.json", false, Encoding.UTF8))
                    {
                        file.Write(sw.ReadToEnd());
                        file.Close();
                    }
                }
            }
            catch (Microsoft.Graph.ServiceException svcEx)
            {
                var additionalData = svcEx.Error.AdditionalData;
                var details = additionalData["details"];
                Console.WriteLine(string.Format("[Error]: '{0}'", details));
                return false;
            }

            #endregion

            #region GetTeamsUserActivityCounts
            try
            {
                // Adding a custom serializer to handle CSV and Transmogrify into JSON (NOTE: check ReportCustomSerializers for specific Report Serializer)
                HttpProvider httpProvider = new HttpProvider(ReportCustomSerializersOptions.teamsUserActivityCountsSerializer);
                GraphServiceClient graphClient = new GraphServiceClient(authProvider, httpProvider);

                // TODO: MS Not implemented this feature on the Reports APIs (Only CSV at this point)
                //var queryOptions = new List<QueryOption>()
                //    {
                //        new QueryOption("$format", "application/json")
                //    };

                var response = await graphClient.Reports.GetTeamsUserActivityCounts(period).Request(/* queryOptions */).GetAsync();
                using (StreamReader sw = new StreamReader(response.Content))
                {
                    using (var file = new StreamWriter("c:\\temp\\GetTeamsUserActivityCounts.json", false, Encoding.UTF8))
                    {
                        file.Write(sw.ReadToEnd());
                        file.Close();
                    }
                }
            }
            catch (Microsoft.Graph.ServiceException svcEx)
            {
                var additionalData = svcEx.Error.AdditionalData;
                var details = additionalData["details"];
                Console.WriteLine(string.Format("[Error]: '{0}'", details));
                return false;
            }

            #endregion

            #region GetTeamsUserActivityCounts
            try
            {
                // Adding a custom serializer to handle CSV and Transmogrify into JSON (NOTE: check ReportCustomSerializers for specific Report Serializer)
                HttpProvider httpProvider = new HttpProvider(ReportCustomSerializersOptions.teamsUserActivityUserDetailSerializer);
                GraphServiceClient graphClient = new GraphServiceClient(authProvider, httpProvider);

                // TODO: MS Not implemented this feature on the Reports APIs (Only CSV at this point)
                //var queryOptions = new List<QueryOption>()
                //    {
                //        new QueryOption("$format", "application/json")
                //    };

                var response = await graphClient.Reports.GetTeamsUserActivityUserDetail(period).Request(/* queryOptions */).GetAsync();
                using (StreamReader sw = new StreamReader(response.Content))
                {
                    using (var file = new StreamWriter("c:\\temp\\GetTeamsUserActivityUserDetail.json", false, Encoding.UTF8))
                    {
                        file.Write(sw.ReadToEnd());
                        file.Close();
                    }
                }
            }
            catch (Microsoft.Graph.ServiceException svcEx)
            {
                var additionalData = svcEx.Error.AdditionalData;
                var details = additionalData["details"];
                Console.WriteLine(string.Format("[Error]: '{0}'", details));
                return false;
            }

            #endregion

            #region GetTeamsDeviceUsageDistributionUserCounts
            try
            {
                // Adding a custom serializer to handle CSV and Transmogrify into JSON (NOTE: check ReportCustomSerializers for specific Report Serializer)
                HttpProvider httpProvider = new HttpProvider(ReportCustomSerializersOptions.teamsDeviceUsageDistributionUserCountsSerializer);
                GraphServiceClient graphClient = new GraphServiceClient(authProvider, httpProvider);

                // TODO: MS Not implemented this feature on the Reports APIs (Only CSV at this point)
                //var queryOptions = new List<QueryOption>()
                //    {
                //        new QueryOption("$format", "application/json")
                //    };

                var response = await graphClient.Reports.GetTeamsDeviceUsageDistributionUserCounts(period).Request(/* queryOptions */).GetAsync();
                using (StreamReader sw = new StreamReader(response.Content))
                {
                    using (var file = new StreamWriter("c:\\temp\\GetTeamsDeviceUsageDistributionUserCounts.json", false, Encoding.UTF8))
                    {
                        file.Write(sw.ReadToEnd());
                        file.Close();
                    }
                }
            }
            catch (Microsoft.Graph.ServiceException svcEx)
            {
                var additionalData = svcEx.Error.AdditionalData;
                var details = additionalData["details"];
                Console.WriteLine(string.Format("[Error]: '{0}'", details));
                return false;
            }

            #endregion

            #region GetTeamsDeviceUsageUserCounts
            try
            {
                // Adding a custom serializer to handle CSV and Transmogrify into JSON (NOTE: check ReportCustomSerializers for specific Report Serializer)
                HttpProvider httpProvider = new HttpProvider(ReportCustomSerializersOptions.teamsDeviceUsageUserCountsSerializer);
                GraphServiceClient graphClient = new GraphServiceClient(authProvider, httpProvider);

                // TODO: MS Not implemented this feature on the Reports APIs (Only CSV at this point)
                //var queryOptions = new List<QueryOption>()
                //    {
                //        new QueryOption("$format", "application/json")
                //    };

                var response = await graphClient.Reports.GetTeamsDeviceUsageUserCounts(period).Request(/* queryOptions */).GetAsync();
                using (StreamReader sw = new StreamReader(response.Content))
                {
                    using (var file = new StreamWriter("c:\\temp\\GetTeamsDeviceUsageUserCounts.json", false, Encoding.UTF8))
                    {
                        file.Write(sw.ReadToEnd());
                        file.Close();
                    }
                }
            }
            catch (Microsoft.Graph.ServiceException svcEx)
            {
                var additionalData = svcEx.Error.AdditionalData;
                var details = additionalData["details"];
                Console.WriteLine(string.Format("[Error]: '{0}'", details));
                return false;
            }

            #endregion

            #region GetTeamsDeviceUsageUserDetail
            try
            {
                // Adding a custom serializer to handle CSV and Transmogrify into JSON (NOTE: check ReportCustomSerializers for specific Report Serializer)
                HttpProvider httpProvider = new HttpProvider(ReportCustomSerializersOptions.teamsDeviceUsageUserDetailSerializer);
                GraphServiceClient graphClient = new GraphServiceClient(authProvider, httpProvider);

                // TODO: MS Not implemented this feature on the Reports APIs (Only CSV at this point)
                //var queryOptions = new List<QueryOption>()
                //    {
                //        new QueryOption("$format", "application/json")
                //    };

                var response = await graphClient.Reports.GetTeamsDeviceUsageUserDetail(period).Request(/* queryOptions */).GetAsync();
                using (StreamReader sw = new StreamReader(response.Content))
                {
                    using (var file = new StreamWriter("c:\\temp\\GetTeamsDeviceUsageUserDetail.json", false, Encoding.UTF8))
                    {
                        file.Write(sw.ReadToEnd());
                        file.Close();
                    }
                }
            }
            catch (Microsoft.Graph.ServiceException svcEx)
            {
                var additionalData = svcEx.Error.AdditionalData;
                var details = additionalData["details"];
                Console.WriteLine(string.Format("[Error]: '{0}'", details));
                return false;
            }

            #endregion


            return true;

        }
    }
}
