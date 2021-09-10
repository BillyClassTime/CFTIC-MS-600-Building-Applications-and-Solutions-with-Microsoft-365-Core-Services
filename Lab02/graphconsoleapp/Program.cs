using System;
using System.Collections.Generic;
using Microsoft.Identity.Client;
using Microsoft.Graph;
using Microsoft.Extensions.Configuration;
using Helpers;

namespace graphconsoleapp
{
    class Program
    {
        private static GraphServiceClient _graphClient;

        static void Main(string[] args)
        {
            Console.WriteLine("Start application");
            var config = LoadAppSettings();
            if(config==null)
            {
                Console.WriteLine("Invalid appsettings.json file.");
                return;
            }

            var client = GetAuthenticatedGraphClient(config);
            var graphRequest = client.Users.Request();
            //var results = graphRequest.GetAsync().Result;
            var results = graphRequest
                          //.Select(u => new {u.Surname,u.JobTitle,u.Id, u.DisplayName, u.Mail })
                          .Select(u => new {u.Surname,u.JobTitle,u.Id, u.Mail })
                          .Top(10)
                          .OrderBy("DisplayName desc")
                          //.Filter("startsWith(surname,'V') or startsWith(surname,'B') or startsWith(surname,'C')")
                          .GetAsync()
                          .Result;
            var i = 0;
            foreach (var user in results)
            {
                //Console.WriteLine($"{i++}:{user.Id}:{user.DisplayName}<{user.Mail}>{user.JobTitle}:Manager:{user.Manager}");
                Console.WriteLine($"{i++}:{user.Id}:QUITADO<{user.Mail}>{user.JobTitle}:Manager:{user.Manager}");
            }
            Console.WriteLine("\nGraph Request:");
            Console.WriteLine(graphRequest.GetHttpRequestMessage().RequestUri);
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
                            string.IsNullOrEmpty(config["applicationSecret"]) ||
                            string.IsNullOrEmpty(config["redirectUri"]) ||
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

        private static IAuthenticationProvider CreateAuthorizationProvider(IConfigurationRoot config)
        {
            var clientId = config["applicationId"];
            var clientSecret = config["applicationSecret"];
            var redirectUri = config["redirectUri"];
            var authority = $"https://login.microsoftonline.com/{config["tenantId"]}/v2.0";
            List<string> scopes = new List<string>();
            scopes.Add("https://graph.microsoft.com/.default");
            var cca = ConfidentialClientApplicationBuilder.Create(clientId)
                                                  .WithAuthority(authority)
                                                  .WithRedirectUri(redirectUri)
                                                  .WithClientSecret(clientSecret)
                                                  .Build();
            return new MsalAuthenticationProvider(cca, scopes.ToArray());
        }

        private static GraphServiceClient GetAuthenticatedGraphClient(IConfigurationRoot config)
        {
            var authenticationProvider = CreateAuthorizationProvider(config);
            _graphClient = new GraphServiceClient(authenticationProvider);
            return _graphClient;
        }

    }
}
