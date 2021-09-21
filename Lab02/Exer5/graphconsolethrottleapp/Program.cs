using System;
using System.Collections.Generic;
using System.Net;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Security;
using System.Threading.Tasks;
using Microsoft.Identity.Client;
using Microsoft.Graph;
using Microsoft.Extensions.Configuration;
using Helpers;
using Newtonsoft.Json;
using static System.Console;
namespace graphconsolethrottleapp
{
    class Program
    {
        static void Main(string[] args)
        {
            WriteLine("Inicio de la applicacion");
            var config = LoadAppSettings();
            if (config == null)
            {
                WriteLine("Invalid apsettings.json file.");
                return;
            }
            var userName = ReadUsername();
            var userPassword = ReadPassword();
            ///
            //var client = GetAuthenticatedHTTPClient(config, userName, userPassword);
            var client = GetAuthenticatedGraphClient(config, userName, userPassword);
            /// Retry strategy
            var stopwatch = new System.Diagnostics.Stopwatch();
            stopwatch.Start();
            /*var clientResponse = client.GetAsync("https://graph.microsoft.com/v1.0/me/messages?$select=id&$top=100").Result;
            // enumerate through the list of messages
            var httpResponseTask = clientResponse.Content.ReadAsStringAsync();
            httpResponseTask.Wait();
            var graphMessages = JsonConvert.DeserializeObject<Messages>(httpResponseTask.Result);
            /// End Retry strategy
            */
            var clientResponse = client.Me.Messages
                                .Request()
                                .Select(m => new { m.Id})
                                .Top(100)
                                .GetAsync()
                                .Result;

            var tasks = new List<Task>();
            //foreach (var graphMessage in graphMessages.Items)
            foreach (var graphMessage in clientResponse.CurrentPage)
            {
                tasks.Add(Task.Run(() =>
                {
                    Console.WriteLine("...retrieving message: {0}", graphMessage.Id);
                    var messageDetail = GetMessageDetail(client, graphMessage.Id);
                    Console.WriteLine("SUBJECT: {0}", messageDetail.Subject);
                }));
            }
            // do all work in parallel & wait for it to complete
            var allWork = Task.WhenAll(tasks);
            try
            {
                allWork.Wait();
            }
            catch { }
            stopwatch.Stop();
            Console.WriteLine();
            Console.WriteLine("Elapsed time: {0} seconds", stopwatch.Elapsed.Seconds);

#region deprecate
            // WriteLine("Después de la authenticación");
            // var totalRequests = 100;
            // var successRequests = 0;
            // var tasks = new List<Task>();
            // var failResponseCode = HttpStatusCode.OK;
            // HttpResponseHeaders failedHeaders = null;
            // WriteLine("Antes de las tareas asyncronas");
            // for (int i = 0; i < totalRequests; i++)
            // {
            //     Write($"{i}-");
            //     tasks.Add(Task.Run(() =>
            //     {
            //         var response = client.GetAsync("https://graph.microsoft.com/v1.0/me/messages").Result;
            //         Write(".");
            //         if (response.StatusCode == HttpStatusCode.OK)
            //         {
            //             successRequests++;
            //         }
            //         else
            //         {
            //             Write('X');
            //             failResponseCode = response.StatusCode;
            //             failedHeaders = response.Headers;
            //         }
            //     }));
            // }
            // WriteLine("Despues del bucle");
            // var allWork = Task.WhenAll(tasks);
            // try
            // {
            //     allWork.Wait();
            // }
            // catch { }
            // WriteLine();
            // WriteLine("{0}/{1} requests succeeded.", successRequests, totalRequests);
            // if (successRequests != totalRequests)
            // {
            //     WriteLine("Failed response code: {0}", failResponseCode.ToString());
            //     WriteLine("Failed response headers: {0}", failedHeaders);
            // }
            // ///
#endregion
        }

        private static IConfigurationRoot LoadAppSettings()
        {
            try
            {
                var config = new ConfigurationBuilder()
                                  .SetBasePath(System.IO.Directory.GetCurrentDirectory())
                                  .AddJsonFile("appsettings.json", false, true)
                                  .Build();
                if (string.IsNullOrEmpty(config["applicationId"]) ||
                    string.IsNullOrEmpty(config["tenantId"]))
                {
                    return null;
                }
                return config;
            }
            catch (System.IO.FileNotFoundException)
            {
                return null;
            }
        }

        private static IAuthenticationProvider CreateAuthorizationProvider(IConfigurationRoot config, string userName, SecureString userPassword)
        {
            var clientId = config["applicationId"];
            var authority = $"https://login.microsoftonline.com/{config["tenantId"]}/v2.0";
            List<string> scopes = new List<string>();
            scopes.Add("User.Read");
            scopes.Add("Mail.Read");
            var cca = PublicClientApplicationBuilder.Create(clientId)
                                                    .WithAuthority(authority)
                                                    .Build();
            return MsalAuthenticationProvider.GetInstance(cca, scopes.ToArray(), userName, userPassword);
        }

        /*private static HttpClient GetAuthenticatedHTTPClient(IConfigurationRoot config, string userName, SecureString userPassword)
        {
            var authenticationProvider = CreateAuthorizationProvider(config, userName, userPassword);
            var httpClient = new HttpClient(new AuthHandler(authenticationProvider, new HttpClientHandler()));
            return httpClient;
        }*/
        private static GraphServiceClient GetAuthenticatedGraphClient(IConfigurationRoot config, string userName, SecureString userPassword)
        {
            var authenticationProvider = CreateAuthorizationProvider(config, userName, userPassword);
            var graphClient = new GraphServiceClient(authenticationProvider);
            return graphClient;
        }


        private static SecureString ReadPassword()
        {
            Console.WriteLine("Enter your password");
            SecureString password = new SecureString();
            while (true)
            {
                ConsoleKeyInfo c = Console.ReadKey(true);
                if (c.Key == ConsoleKey.Enter)
                {
                    break;
                }
                password.AppendChar(c.KeyChar);
                Console.Write("*");
            }
            Console.WriteLine();
            return password;
        }

        private static string ReadUsername()
        {
            string username;
            Console.WriteLine("Enter your username");
            username = Console.ReadLine();
            return username;
        }

        private static Message GetMessageDetail(GraphServiceClient client, string messageId)
        {
            return client.Me.Messages[messageId].Request().GetAsync().Result;
        }

        private static Message GetMessageDetail(HttpClient client, string messageId, int defaultDelay = 2)
        {
            Message messageDetail = null;
            string endpoint = "https://graph.microsoft.com/v1.0/me/messages/" + messageId;

            // submit request to Microsoft Graph & wait to process response
            var clientResponse = client.GetAsync(endpoint).Result;
            var httpResponseTask = clientResponse.Content.ReadAsStringAsync();
            httpResponseTask.Wait();

            WriteLine("...Response status code: {0}  ", clientResponse.StatusCode);
            // IF request successful (not throttled), set message to retrieved message
            if (clientResponse.StatusCode == HttpStatusCode.OK)
            {
                messageDetail = JsonConvert.DeserializeObject<Message>(httpResponseTask.Result);
            }
            // ELSE IF request was throttled (429, aka: TooManyRequests)...
            else if (clientResponse.StatusCode == HttpStatusCode.TooManyRequests)
            {
                // get retry-after if provided; if not provided default to 2s
                int retryAfterDelay = defaultDelay;
                if (clientResponse.Headers.RetryAfter.Delta.HasValue && (clientResponse.Headers.RetryAfter.Delta.Value.Seconds > 0))
                {
                    retryAfterDelay = clientResponse.Headers.RetryAfter.Delta.Value.Seconds;
                }
                // wait for specified time as instructed by Microsoft Graph's Retry-After header,
                //    or fall back to default
                WriteLine(">>>>>>>>>>>>> sleeping for {0} seconds...", retryAfterDelay);
                System.Threading.Thread.Sleep(retryAfterDelay * 1000);
                // call method again after waiting
                messageDetail = GetMessageDetail(client, messageId);
            }


            // add code here
            return messageDetail;
        }



    }
}
