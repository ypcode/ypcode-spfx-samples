using Microsoft.Graph;
using Microsoft.Identity.Client;
using MicrosoftGraphWebApp.TokenStorage;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Net.Http.Headers;
using System.Security.Claims;
using System.Threading.Tasks;
using System.Web;

namespace MicrosoftGraphWebApp.Helpers
{
    public static class GraphHelper
    {

        // Load configuration settings from PrivateSettings.config
        private static string appId = ConfigurationManager.AppSettings["AppId"];
        private static string appSecret = ConfigurationManager.AppSettings["AppSecret"];
        private static string redirectUri = ConfigurationManager.AppSettings["RedirectUri"];
        private static string graphScopes = ConfigurationManager.AppSettings["AppScopes"];

        public static async Task<IEnumerable<Event>> GetEventsAsync()
        {
            var graphClient = GetAuthenticatedClient();

            var events = await graphClient.Me.Events.Request()
                .Select("subject,organizer,start,end")
                .OrderBy("createdDateTime DESC")
                .GetAsync();

            return events.CurrentPage;
        }

        public static async Task<IEnumerable<Message>> GetMyEmails()
        {
            var graphClient = GetAuthenticatedClient();

            var messages = await graphClient.Me.Messages.Request().Filter("isRead eq false").Top(3).GetAsync();

            return messages;
        }

        public static async Task<IEnumerable<DriveItem>> GetMyFiles(string path = null)
        {
            var graphClient = GetAuthenticatedClient();

            var query = string.IsNullOrEmpty(path)
                ? graphClient.Me.Drive.Root
                : graphClient.Me.Drive.Root.ItemWithPath(path);

            var files = await query.Children.Request().GetAsync();

            return files;
        }

        private static GraphServiceClient GetAuthenticatedClient()
        {
            return new GraphServiceClient(
                new DelegateAuthenticationProvider(
                    async (requestMessage) =>
                    {
                        // Get the signed in user's id and create a token cache
                        string signedInUserId = ClaimsPrincipal.Current?.FindFirst(ClaimTypes.NameIdentifier)?.Value;
                        if (string.IsNullOrEmpty(signedInUserId))
                            return;

                        SessionTokenStore tokenStore = new SessionTokenStore(signedInUserId,
                            new HttpContextWrapper(HttpContext.Current));

                        var idClient = new ConfidentialClientApplication(
                            appId, redirectUri, new ClientCredential(appSecret),
                            tokenStore.GetMsalCacheInstance(), null);

                        var accounts = await idClient.GetAccountsAsync();

                        // By calling this here, the token can be refreshed
                        // if it's expired right before the Graph call is made
                        var result = await idClient.AcquireTokenSilentAsync(
                                    graphScopes.Split(' '), accounts.FirstOrDefault());

                        requestMessage.Headers.Authorization =
                            new AuthenticationHeaderValue("Bearer", result.AccessToken);
                    }));
        }

        public static async Task<User> GetUserDetailsAsync(string accessToken)
        {
            if (string.IsNullOrEmpty(accessToken))
                return null;

            var graphClient = new GraphServiceClient(
                new DelegateAuthenticationProvider(
                    async (requestMessage) =>
                    {
                        requestMessage.Headers.Authorization =
                            new AuthenticationHeaderValue("Bearer", accessToken);
                    }));

            return await graphClient.Me.Request().GetAsync();
        }
    }
}