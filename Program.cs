using AccessTeamsReports.Utilities;
using Azure.Core;
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
           
            string[] scopes = new string[] { "UserAuthenticationMethod.ReadWrite.All", "User.Read.All" };
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
                const string AAD_USER = "marframfa@corusdev.onmicrosoft.com";
                GraphServiceClient graphClient = new GraphServiceClient(authProvider);

                // Step 1. Get the User ID for the given UPN
                #region Step 1
                    var queryOptionsUsers = new List<QueryOption>()
                                {
                                    new QueryOption("$filter", string.Format("userPrincipalName eq '{0}'",AAD_USER)),
                                    new QueryOption("$format", "application/json")
                                };

                var responseUserInfo = await graphClient.Users.Request(queryOptionsUsers).GetAsync();

                string aadUserID = responseUserInfo[0].Id;

                #endregion

                // Step 2. Get the ID (Record) of the phoneAuthenticationMethod that we want (mobile in this case)
                #region Step 2
                var queryOptionsAuthMethod = new List<QueryOption>()
                            {
                                // Options are 'mobile', 'alternateMobile', 'office'
                                new QueryOption("$filter", "phoneType eq 'mobile'"),
                                new QueryOption("$format", "application/json")
                            };
                    var responseAuthMethod = await graphClient.Users[aadUserID]
                                         .Authentication.PhoneMethods.Request(queryOptionsAuthMethod).GetAsync();
                #endregion

                // Step 3. Get the ID of the AuthenticationMethod type (Mobile) for the next Graph Call
                #region Step 3
                    string AuthMethodId = responseAuthMethod[0].Id;
                #endregion

                // Step 4. Update the Mobile Phone based on the Phone Authentication Method
                #region step 4
                    var phoneAuthenticationMethod = new PhoneAuthenticationMethod
                    {
                        PhoneNumber = "+1 2065550000",
                        PhoneType = AuthenticationPhoneType.Mobile
                    };

                    // Serialize the PhoneAuthenticationMethod to pass to the low level HTTP functions as we are 
                    // bypassing the Graph Client SDK for this one call
                    string jsonphoneAuthenticationMethod = System.Text.Json.JsonSerializer.Serialize(phoneAuthenticationMethod, 
                           new JsonSerializerOptions { WriteIndented = true, IgnoreNullValues = true });

                // NOTE: THIS SDK CODE IS BROKEN Because it sends a PATCH not a PUT
                //var responsePut = await graphClient.Users[AAD_USER_ID]
                //                    .Authentication.PhoneMethods[PhoneMethodID].Request().UpdateAsync(phoneAuthenticationMethod);
                var formulatedGraphRequestRequest = graphClient.Users[aadUserID]
                                    .Authentication.PhoneMethods[AuthMethodId].Request().GetHttpRequestMessage();

                // Update the body + the Method to PUT as UpdateAsynch uses Patch
                formulatedGraphRequestRequest.Content = new StringContent(jsonphoneAuthenticationMethod, Encoding.UTF8, "application/json");
                formulatedGraphRequestRequest.Method = HttpMethod.Put;

                // Call the GraphClientFactory to get the raw HTTP request plus Authentication + 
                // Rate limiting capabilities, etc, which are already built into the request object with the SDK
                var httpGraphConfiguredClient = GraphClientFactory.Create(authProvider);

                var responseAuthMethodMobileUpdate = httpGraphConfiguredClient.SendAsync(formulatedGraphRequestRequest).GetAwaiter().GetResult();

                if (responseAuthMethodMobileUpdate.StatusCode != System.Net.HttpStatusCode.OK)
                {
                    return false;
                }

                #endregion

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
