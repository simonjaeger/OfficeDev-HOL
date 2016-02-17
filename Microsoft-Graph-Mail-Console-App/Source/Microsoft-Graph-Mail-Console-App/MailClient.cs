using Microsoft.IdentityModel.Clients.ActiveDirectory;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
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
            get { return "simonj.onmicrosoft.com"; }
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
            get { return "b2daa454-f3ea-46b4-aeb7-305b109e2a1c"; }
        }

        // The redirect URI specified in the Azure AD application
        public static Uri RedirectUri
        {
            get { return new Uri("https://l0"); }
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

        // Sends a mail to the current user
        public static async Task SendMeAsync()
        {
            // Create the authentication context (ADAL)
            var authenticationContext = new AuthenticationContext(Authority);

            // Get the access token
            var authenticationResult = authenticationContext.AcquireToken(GraphResource,
                ClientId, RedirectUri, PromptBehavior.Auto);
            var accessToken = authenticationResult.AccessToken;

            // Create the HTTP client with the access token
            var httpClient = new HttpClient();
            httpClient.DefaultRequestHeaders.Authorization =
                new AuthenticationHeaderValue("Bearer",
                accessToken);

            // Get and deserialize the user
            var userResponse = await httpClient.GetStringAsync(GraphResource + GraphVersion + "/me/");
            var user = JsonConvert.DeserializeObject<UserModel>(userResponse);

            // Create the mail
            var sendMailRequest = new SendMailRequestModel
            {
                Message = new MessageModel
                {
                    ToRecipients = new List<ToRecipientModel>
                    {
                        new ToRecipientModel
                        {
                            EmailAddress = new EmailAddressModel
                            {
                                Address = user.Mail
                            }
                        }
                    },
                    Subject = "Hello #Office365Dev",
                    Body = new BodyModel
                    {
                        ContentType = "html",
                        Content = "<strong>Lorem ipsum dolor sit amet</strong>, consectetur adipiscing " +
                        "elit, sed do eiusmod tempor incididunt ut labore et dolore " +
                        "magna aliqua. Ut enim ad minim veniam, quis nostrud " +
                        "exercitation ullamco laboris nisi ut aliquip ex ea commodo " +
                        "consequat. Duis aute irure dolor in reprehenderit in voluptate " +
                        "velit esse cillum dolore eu fugiat nulla pariatur. Excepteur sint " +
                        "occaecat cupidatat non proident, sunt in culpa qui officia deserunt " +
                        "mollit anim id est laborum."
                    }
                }
            };

            // Send the mail
            var stringContent = JsonConvert.SerializeObject(sendMailRequest);
            var response = await httpClient.PostAsync(GraphResource + GraphVersion + "/me/microsoft.graph.sendMail",
                new StringContent(stringContent, Encoding.UTF8, "application/json"));

            if (response.IsSuccessStatusCode)
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
