using AccessTeamsReports.Utilities;
using Microsoft.Graph;
using Microsoft.Graph.Auth;
using Microsoft.Identity.Client;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Text;
using System.Text.Json;
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

        public static string accessToken = string.Empty;
   
        static async Task Main(string[] args)
        {

            #region Authentication work Here
            IPublicClientApplication publicClientApplication = PublicClientApplicationBuilder
                        .Create(ConfigurationManager.AppSettings["APP_CLIENT_ID"])
                        .WithAuthority(ConfigurationManager.AppSettings["AAD_AUTHORITY"])
                        .WithDefaultRedirectUri()
                        .Build();
           
            string[] scopes = new string[] { "UserAuthenticationMethod.ReadWrite.All", "Organization.ReadWrite.All","Reports.Read.All" };
            var accounts = await publicClientApplication.GetAccountsAsync();
            
            AuthenticationResult result;
            try
            {
                result = await publicClientApplication.AcquireTokenSilent(scopes, accounts.FirstOrDefault())
                            .ExecuteAsync();
                
                accessToken = result.AccessToken;
            }
            catch (MsalUiRequiredException)
            {
                result = await publicClientApplication.AcquireTokenInteractive(scopes)
                            .ExecuteAsync();

                accessToken = result.AccessToken;
            }

            // Create the Autentication Provider with Proper Scopes for Graph APIs
            DeviceCodeProvider authProvider = new DeviceCodeProvider(publicClientApplication, scopes);

            #endregion

            
            #region Calling Graph Beta APIs
            // Call the APIs with the newly formed Authentication Provider
            bool success = await RunAuthenticationMethodPhoneUpdate(authProvider);

            if (success)
                Console.WriteLine("Jobs Finished!");
            else
                Console.WriteLine("Error!");

            #endregion

            Console.WriteLine("Press <ENTER> to close Application");
            Console.Read();


        }

        private async static Task<bool> RunAuthenticationMethodPhoneUpdate(DeviceCodeProvider authProvider)
        {
           
            #region AuthenticationMethodPhoneUpdate
            try
            {
                const string AAD_USER_ID = "2fcce8ce-dd74-4e86-afec-66733b59a06e";
                GraphServiceClient graphClient = new GraphServiceClient(authProvider);
                
                var queryOptions = new List<QueryOption>()
                    {
                        // Options are 'mobile', 'alternateMobile', 'office'
                        new QueryOption("$filter", "phoneType eq 'mobile'"),
                        new QueryOption("$format", "application/json")
                    };

                // Get the ID (Record) of the phoneAuthenticationMethod that we want 
                var responseGet = await graphClient.Users[AAD_USER_ID]
                                     .Authentication.PhoneMethods.Request(queryOptions).GetAsync();

                // The Phone Update Method to udate based on ID
                string PhoneMethodID = responseGet[0].Id;
                
                // Update the Mobile Phone based on the Phone Authentication Method Above
                var phoneAuthenticationMethod = new PhoneAuthenticationMethod
                {
                    PhoneNumber = "+1 2065550000",
                    PhoneType = AuthenticationPhoneType.Mobile
                };

                // Serialize the PhoneAuthenticationMethod to pass to the low level HTTP functions as we are 
                // bypassing the Graph Client SDK for this one call
                string jsonphoneAuthenticationMethod = System.Text.Json.JsonSerializer.Serialize(phoneAuthenticationMethod, new JsonSerializerOptions { WriteIndented = true, IgnoreNullValues = true });

                // NOTE: THIS SDK CODE IS BROKEN Because it sends a PATCH not a PUT
                //var responsePut = await graphClient.Users[AAD_USER_ID]
                //                    .Authentication.PhoneMethods[PhoneMethodID].Request().(phoneAuthenticationMethod);

                // Manually Constructing the PUT Call to the Graph API
                HttpRequestMessage request = new HttpRequestMessage(HttpMethod.Put,
                    string.Format("https://graph.microsoft.com/beta/users/{0}/authentication/phoneMethods/{1}",
                                    AAD_USER_ID, PhoneMethodID));
                request.Content = new StringContent(jsonphoneAuthenticationMethod, Encoding.UTF8, "application/json");
                request.Headers.Add("Authorization", string.Format("Bearer {0}", accessToken));

                var responsePut = await graphClient.HttpProvider.SendAsync(request);

                if (responsePut.StatusCode != System.Net.HttpStatusCode.OK)
                {
                    return false;
                }
            }
            catch (Exception svcEx)
            {
                
                Console.WriteLine(string.Format("[Error]: '{0}'", svcEx.Message));
                return false;
            }

            return true;

            #endregion
                      
           

        }
    }
}
