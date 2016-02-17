using Microsoft.IdentityModel.Clients.ActiveDirectory;
using Newtonsoft.Json;
using System;
using System.Linq;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Text;
using System.Threading.Tasks;

namespace Microsoft_Graph_Mail_Console_App
{
    public static class MailClient
    {
        // The Azure AD instance where you domain is hosted
        public static string AADInstance
        {
            get { return "https://login.microsoftonline.com"; }
        }

        // The Office 365 domain (e.g. contoso.microsoft.com)
        public static string Domain
        {
            get { return "[YOUR DOMAIN]"; }
        }

        // The authority for authentication; combining the AADInstance
        // and the domain.
        public static string Authority
        {
            get { return string.Format("{0}/{1}/", AADInstance, Domain); }
        }

        // The client ID of your native Azure AD application
        public static string ClientId
        {
            get { return "[YOUR CLIENT ID]"; }
        }

        // The redirect URI specified in the Azure AD application
        public static Uri RedirectUri
        {
            get { return new Uri("[YOUR REDIRECT URI]"); }
        }

        // The resource identifier for the Microsoft Graph
        public static string GraphResource
        {
            get { return "https://graph.microsoft.com/"; }
        }

        // The Microsoft Graph version, can be "v1.0" or "beta"
        public static string GraphVersion
        {
            get { return "v1.0"; }
        }

        // Get an access token for the Microsoft Graph using ADAL
        public static string GetAccessToken()
        {
            // Create the authentication context (ADAL)
            var authenticationContext = new AuthenticationContext(Authority);

            // Get the access token
            var authenticationResult = authenticationContext.AcquireToken(GraphResource,
                ClientId, RedirectUri, PromptBehavior.RefreshSession);
            var accessToken = authenticationResult.AccessToken;
            return accessToken;
        }

        // Prepare an HttpClient with the an authorization header (access token)
        public static HttpClient GetHttpClient(string accessToken)
        {
            // Create the HTTP client with the access token
            var httpClient = new HttpClient();
            httpClient.DefaultRequestHeaders.Authorization =
                new AuthenticationHeaderValue("Bearer",
                accessToken);
            return httpClient;
        }

        // Get the current user using a prepared HttpClient (with an authorization header)
        public static async Task<UserModel> GetUserAsync(HttpClient httpClient)
        {
            // Get and deserialize the user
            var userResponse = await httpClient.GetStringAsync(GraphResource + GraphVersion + "/me/");
            var user = JsonConvert.DeserializeObject<UserModel>(userResponse);
            return user;
        }

        // Create the mail structure needed to make a request to the Microsoft Graph
        public static MailModel CreateMail(string subject, string body, params string[] recipients)
        {
            var mail = new MailModel
            {
                Message = new MessageModel
                {
                    ToRecipients = recipients.Select(recipient => new ToRecipientModel
                    {
                        EmailAddress = new EmailAddressModel
                        {
                            Address = recipient
                        }
                    }).ToList(),
                    Subject = subject,
                    Body = new BodyModel
                    {
                        ContentType = "html",
                        Content = body
                    }
                }
            };
            return mail;
        }

        // Send a mail using a prepared HttpClient (with an authorization header)
        private static async Task<bool> SendMailAsync(HttpClient httpClient, MailModel mail)
        {
            // Serialize the mail
            var stringContent = JsonConvert.SerializeObject(mail);

            // Send using POST
            var response = await httpClient.PostAsync(GraphResource + GraphVersion + "/me/microsoft.graph.sendMail",
                new StringContent(stringContent, Encoding.UTF8, "application/json"));
            return response.IsSuccessStatusCode;
        }

        // Send a mail to the current user
        public static async Task SendMeAsync()
        {
            // Get an access token and configure the HttpClient
            var accessToken = GetAccessToken();
            var httpClient = GetHttpClient(accessToken);

            // Get the current user (to extract the mail address)
            var user = await GetUserAsync(httpClient);

            // Create the mail (HTML)
            var mail = CreateMail("Hello #Office365Dev",
                "<strong>Lorem ipsum dolor sit amet</strong>, consectetur adipiscing " +
                "elit, sed do eiusmod tempor incididunt ut labore et dolore " +
                "magna aliqua. Ut enim ad minim veniam, quis nostrud " +
                "exercitation ullamco laboris nisi ut aliquip ex ea commodo " +
                "consequat. Duis aute irure dolor in reprehenderit in voluptate " +
                "velit esse cillum dolore eu fugiat nulla pariatur. Excepteur sint " +
                "occaecat cupidatat non proident, sunt in culpa qui officia deserunt " +
                "mollit anim id est laborum.", user.Mail);

            // Send the mail and check for success
            var isSuccess = await SendMailAsync(httpClient, mail);
            if (isSuccess)
            {
                Console.Write("Ooh... check your mailbox!");
            }
            else
            {
                // TODO: Handle error
                Console.Write("Oops... something went wrong!");
            }
        }
    }
}
