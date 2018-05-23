using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Azure.ActiveDirectory.GraphClient;
using Microsoft.IdentityModel.Clients.ActiveDirectory;
using System.Net.Http.Headers;
using System.Globalization;
using System.Net.Http;
using Newtonsoft.Json;
using System.Diagnostics;
using System.IO;

namespace sharepoint_graph
{
    class Program
    {
        //note native applications can run either in user mode or app mode
        //in user mode the user needs to authenticate
        //in application mode no user interaction is needed for authentication and accessing resources
        public static string TokenForUser;
        public static string GraphTokenForUser;
        public static string GraphTokenForApplication;

        //public const string AuthString = "https://login.microsoftonline.com/qwcdev.onmicrosoft.com";
        //public const string AuthString = "https://login.microsoftonline.com/common/oauth2/v2.0/token"; //o365 account
        public const string AuthString = "https://login.microsoftonline.com/hydroexperiment.onmicrosoft.com"; //o365 account
        public const string ResourceUrl = "https://graph.windows.net";
        public const string GraphResourceUrl = "https://graph.microsoft.com/";
        public const string GraphServiceObjectId = "00000002-0000-0000-c000-000000000000";
        //public const string TenantId = "cd05c038-6b46-45ab-9948-e4c7a614ec4a"; //qwcdev
        //public const string TenantId = "06ec30ca-c091-495d-9730-8460f2ab57d8"; //trial o365 site
        public const string TenantId = "412c8639-9818-4b46-a7d5-943d5c917aef"; //hydroexp
        //public const string ClientId = "58306735-0a0d-441b-92e3-3d2541ae22f5"; //clientID for User Mode Registration
        //public const string ClientId = "dae8a81e-1583-448e-8ace-93513a27478a"; //clientID for app mode registration
        //public const string ClientId = "654f6548-fe3a-4616-b2dd-ca908ebf8ca7"; //qwcdev app mode registration
        //public const string ClientId = "763408b5-9479-4782-b399-9a4662324ed7"; //trail o365 site
        public const string ClientId = "75cb3937-fce6-468b-a4be-d1c04ce5b8bd"; //batchgrpError calling graph API Cannot send a content-body with this verb-type.
        //used for app mode operation
        //public const string ClientSecret = "cWhtNLTYphHmY5oVsMKhjIjQ5ARdfrzZGZ/TyF/uEWE="; //qwcdev
        //public const string ClientSecret = "Ir6LDHT1/aIEA4/8+99Zd6DhAr0vHaZ24D+QyV5MqCY="; //o365 trial site
        public const string ClientSecret = "Pe24VjubuEduJRYRVnwjs6rJIumH4Hzw8uNTXgUR3nU=";//batch

        static void Main(string[] args)
        {
            //string x = GetTokenForUser().Result;
            GetGraphTokenForApplication().Wait();
            var result = CallGraph2().Result;
        }

        public static async Task<string> GetTokenForUser()
        {
            if (TokenForUser == null)
            {
                var redirectUri = new Uri("https://www.google.com/");
                AuthenticationContext authenticationContext = new AuthenticationContext(AuthString, false);
                AuthenticationResult userAuthnResult = await authenticationContext.AcquireTokenAsync(ResourceUrl,
                    ClientId, redirectUri, new PlatformParameters(PromptBehavior.RefreshSession));
                TokenForUser = userAuthnResult.AccessToken;
                Console.WriteLine("\n Welcome " + userAuthnResult.UserInfo.GivenName + " " +
                                  userAuthnResult.UserInfo.FamilyName);
            }
            return TokenForUser;
        }

        //public static async Task<string> GetGraphTokenForUser()
        //{
        //    //to get token for graph resource 
        //    if (GraphTokenForUser == null)
        //    {
        //        var redirectUri = new Uri("https://localhost");
        //        AuthenticationContext authenticationContext = new AuthenticationContext(AuthString, false);
        //        AuthenticationResult userAuthnResult = await authenticationContext.AcquireTokenAsync(GraphResourceUrl,
        //            ClientId, redirectUri, new PlatformParameters(PromptBehavior.RefreshSession));
        //        GraphTokenForUser = userAuthnResult.AccessToken;
        //        Console.WriteLine("\n Welcome " + userAuthnResult.UserInfo.GivenName + " " +
        //                          userAuthnResult.UserInfo.FamilyName);
        //    }
        //    return GraphTokenForUser;
        //}

        public static async Task<string> GetGraphTokenForApplication()
        {
            //string GraphTokenForApplication = null;
            if (GraphTokenForApplication == null)
            {
                AuthenticationContext authenticationContext = new AuthenticationContext(AuthString, false);

                // Configuration for OAuth client credentials 
                if (string.IsNullOrEmpty(ClientSecret))
                {
                    Console.WriteLine("Client secret not set. Please follow the steps in the README to generate a client secret.");
                }
                else
                {
                    ClientCredential clientCred = new ClientCredential(ClientId, ClientSecret);
                    AuthenticationResult authenticationResult = await authenticationContext.AcquireTokenAsync(GraphResourceUrl, clientCred);
                    GraphTokenForApplication = authenticationResult.AccessToken;
                }
            }
            return GraphTokenForApplication;
        }

        public static ActiveDirectoryClient GetActiveDirectoryClientAsUser()
        {
            Uri servicePointUri = new Uri(ResourceUrl);
            Uri serviceRoot = new Uri(servicePointUri, TenantId);
            ActiveDirectoryClient activeDirectoryClient = new ActiveDirectoryClient(serviceRoot,
                async () => await GetGraphTokenForApplication());
            return activeDirectoryClient;
        }

        public static async Task<User> GetSignedInUser(ActiveDirectoryClient client)
        {
            User signedInUser = new User();
            try
            {
                signedInUser = (User)await client.Me.ExecuteAsync();
                Console.WriteLine("\nUser UPN: {0}, DisplayName: {1}", signedInUser.UserPrincipalName,
                    signedInUser.DisplayName);

            }
            catch (Exception e)
            {
                Console.WriteLine("\nError getting signed in user {0}", e.InnerException);
            }
            return signedInUser;

        }

        public static async Task<string> CallGraph(string url)
        {
            //string callResult = "done";
            try
            {
                //GetGraphTokenForUser().Wait();
                //string tok = await GetGraphTokenForApplication();

                string requestUrl = String.Format(
                                       CultureInfo.InvariantCulture,
                                       url,
                                       TenantId);
                HttpClient x = new HttpClient();
                HttpRequestMessage request = new HttpRequestMessage(HttpMethod.Post, requestUrl);
                request.Headers.Authorization = new AuthenticationHeaderValue("Bearer", GraphTokenForApplication);


                HttpResponseMessage response = await x.SendAsync(request);

                HttpContent content = response.Content;
                string body = await content.ReadAsStringAsync();

                //Console.WriteLine(body);
                return body;

            }
            catch (Exception e)
            {
                Console.WriteLine("\nError calling graph API {0}", e.Message);
                Console.ReadKey();
                return "error";
            }


        }

        public static async Task<string> CallGraph2()
        {
            string url = "https://graph.microsoft.com/v1.0/$batch";
            //string body = ""
            //string body = " { \"requests\": [ { \"url\": \" /\", \"method\": \"GET\", \"id\": \"1\" }, { \"url\": \" / me / messages ?$filter = importance eq 'high' &$select = from,subject,receivedDateTime,bodyPreview\", \"method\": \"GET\", \"id\": \"2\" }, { \"url\": \" / me / events\", \"method\": \"GET\", \"id\": \"3\" } ] } ";
            //string bodys = "{\"requests\": [{\"url\": \"/me?$select=displayName,jobTitle,userPrincipalName\",\"method\": \"GET\",\"id\": \"1\"},{\"url\": \"/me/messages?$filter=importance eq \'high\'&$select=from,subject,receivedDateTime,bodyPreview\",\"method\": \"GET\",\"id\": \"2\"},{\"url\": \"/me/events\",\"method\": \"GET\",\"id\": \"3\"}]}";
            string bodys = " ";

            //string callResult = "done";
            try
            {
                //GetGraphTokenForUser().Wait();
                //string tok = await GetGraphTokenForApplication();

                string requestUrl = String.Format(
                                       CultureInfo.InvariantCulture,
                                       url,
                                       TenantId);
                HttpClient x = new HttpClient();
                HttpRequestMessage request = new HttpRequestMessage(HttpMethod.Post, requestUrl);
                request.Headers.Authorization = new AuthenticationHeaderValue("Bearer", GraphTokenForApplication);
                ////request.Content = new StringContent(body,
                //                    Encoding.UTF8,
                //                    "application/json");


                HttpResponseMessage response = await x.SendAsync(request);

                HttpContent content = response.Content;
                string resp = await content.ReadAsStringAsync();

                //Console.WriteLine(body);
                return resp;

            }
            catch (Exception e)
            {
                Console.WriteLine("\nError calling graph API {0}", e.Message);
                Console.ReadKey();
                return "error";
            }


        }
    }
}
