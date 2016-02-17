# Hands-on Lab: Microsoft Graph Mail Console App #

If you haven’t heard, there is an easy way to call a great amount of Microsoft APIs – using one single endpoint. This endpoint, so called the Microsoft Graph (https://graph.microsoft.io/) lets you access everything from data, to intelligence and insights powered by the Microsoft cloud.

No longer will you need to keep track of different endpoints and separate tokens in your solutions – how great is that? This post is an introductory part of getting started with the Microsoft Graph. For changes in the Microsoft Graph, head to: https://graph.microsoft.io/changelog

This hands-on lab demonstrates how to configure Azure AD to acquire an access token, used when calling the Microsoft Graph. With the Microsoft Graph, we will send ourselves a mail from the console application.
The objective is to get familiar with the authentication piece (which is the complex part), but also to display some of the possibilities of the Microsoft Graph.

The hands-on lab is divided into multiple exercises and should be followed in a chronological order. These are the included exercises:

* [1.1 Register the application in Azure AD](#exercise-11-register-the-application-in-azure-ad)
* [1.2 Configure the application in Azure AD](#exercise-12-configure-the-application-in-azure-ad)
* [2.1 Create the project](#exercise-21-create-the-project)
* [2.2 Add the dependencies](#exercise-22-add-the-dependencies)
* [3.1 Add the models](#exercise-31-add-the-models)
* [3.2 Add the MailClient class](#exercise-32-add-the-mailclient-class)
* [3.3 Configure the MailClient class](#exercise-33-configure-the-mailclient-class)
* [3.4 Finish the MailClient class](#exercise-34-finish-the-mailclient-class)
* [3.5 Finish the Program class](#exercise-35-finish-the-program-class)

Short of time and just want the final sample? Clone this repository (```git clone https://github.com/simonjaeger/OfficeDev-HOL.git```) and open the solution file: **Microsoft-Graph-Mail-Console-App\\Source\\Microsoft-Graph-Mail-Console-App.sln**.           
    
### Prerequisites ###
- Visual Studio 2015: <https://www.visualstudio.com/en-us/downloads/download-visual-studio-vs.aspx>
- Office 365 Developer Tenant: <http://dev.office.com/devprogram>
- Azure subscription for the Office 365 Developer Tenant (a trial is sufficient)

### Solution ###
Solution | Author(s)
---------|----------
Microsoft-Graph-Mail-Console-App | Simon Jäger (**Microsoft**)

### Version history ###
Version  | Date | Comments
---------| -----| --------
1.1  | February 17th 2016 | Code restructure
1.0  | February 16th 2016 | Initial release

### Disclaimer ###
**THIS CODE IS PROVIDED *AS IS* WITHOUT WARRANTY OF ANY KIND, EITHER EXPRESS OR IMPLIED, INCLUDING ANY IMPLIED WARRANTIES OF FITNESS FOR A PARTICULAR PURPOSE, MERCHANTABILITY, OR NON-INFRINGEMENT.**

----------

# Exercises #
#### Exercise 1.1: Register the application in Azure AD ####
The first thing that we need to do is to register our application in Azure AD. This is how can acquire an access token for the Microsoft Graph (on behalf of the user). 

Make sure that you have all of the required prerequisites before getting into the exercise.

1. Head to the current Azure portal and sign in with a user in your Office 365 Developer Tenant at: <https://manage.windowsazure.com/>
2. In the left menu, choose on **Active Directory** to list your Azure AD.  
   ![](https://raw.githubusercontent.com/simonjaeger/OfficeDev-HOL/master/Microsoft-Graph-Mail-Console-App/Images/AAD.png)

3. Click on your Azure AD in the **Directory** tab section.  
   ![](https://raw.githubusercontent.com/simonjaeger/OfficeDev-HOL/master/Microsoft-Graph-Mail-Console-App/Images/SelectTenant.png)
4. Click on **Applications** in the tab menu. 
5. Click on the **Add** button at the bottom to register a new Azure AD application.   
   ![](https://raw.githubusercontent.com/simonjaeger/OfficeDev-HOL/master/Microsoft-Graph-Mail-Console-App/Images/AddApp.png)
6. Choose on **Add an application my organization is developing** in the dialog that shows up.
7. Name your application **"Microsoft-Graph-Mail-Console-App"** and select **Native Client Application**. Proceed by clicking the continue button down at the bottom.   
   ![](https://raw.githubusercontent.com/simonjaeger/OfficeDev-HOL/master/Microsoft-Graph-Mail-Console-App/Images/AddApp2.png)
8. Enter **"https://Microsoft-Graph-Mail-Console-App/"** as the **Redirect URI**. Click on the check button down below to complete the wizard.   
   ![](https://raw.githubusercontent.com/simonjaeger/OfficeDev-HOL/master/Microsoft-Graph-Mail-Console-App/Images/AddApp3.png)


#### Exercise 1.2: Configure the application in Azure AD ####
We need to configure the application in Azure AD to be able to request the Microsoft Graph as a resource. This is done by adding the Microsoft Graph as an application, along with a set of permissions.

1. Click on the **Configure** tab in the Azure AD application page.   
   ![](https://raw.githubusercontent.com/simonjaeger/OfficeDev-HOL/master/Microsoft-Graph-Mail-Console-App/Images/ConfigureApp.png)
2. Scroll down to the bottom to the **Permissions to other applications** section and click on the **Add application** button.
3. Show **Microsoft Apps** and add the **Microsoft Graph**. 
4. Save your changes by clicking the check button at the bottom.   
   ![](https://raw.githubusercontent.com/simonjaeger/OfficeDev-HOL/master/Microsoft-Graph-Mail-Console-App/Images/AddMSGraph.png)
5. Click on **Delegated permissions** and pick the **Send mail as a user** permission.   
   ![](https://raw.githubusercontent.com/simonjaeger/OfficeDev-HOL/master/Microsoft-Graph-Mail-Console-App/Images/AddPermissions.png)
6. Save your configuration by clicking the **Save** button at the bottom.   
   ![](https://raw.githubusercontent.com/simonjaeger/OfficeDev-HOL/master/Microsoft-Graph-Mail-Console-App/Images/SaveApp.png)

#### Exercise 2.1: Create the project ####
Now let's create the project itself in Visual Studio 2015.

1. Click **File**, **New** and finally the **Project** button.
2. In **Templates**, select **Visual C#**, **Windows** and then choose **Console Application**. 
3. Name your project **"Microsoft-Graph-Mail-Console-App"** and click the **OK** button to continue. 
   ![](https://raw.githubusercontent.com/simonjaeger/OfficeDev-HOL/master/Microsoft-Graph-Mail-Console-App/Images/NewProject.png)
4. Visual Studio 2015 will now create your console application project. Have a look in the **Solution Explorer** and find your new project.  
   ![](https://raw.githubusercontent.com/simonjaeger/OfficeDev-HOL/master/Microsoft-Graph-Mail-Console-App/Images/SolutionExplorer.png)

#### Exercise 2.2: Add the dependencies ####
We will use a couple of NuGet packages in this console application, to make it easy for us. These are Json.NET (<http://www.newtonsoft.com/json>) and Active Directory Authentication Library (<https://azure.microsoft.com/sv-se/documentation/articles/active-directory-authentication-libraries/>).

1. Click on **Tools** in the Visual Studio 2015 top menu.
2. Select **NuGet Package Manager** and choose **Package Manager Console**.  
   ![](https://raw.githubusercontent.com/simonjaeger/OfficeDev-HOL/master/Microsoft-Graph-Mail-Console-App/Images/PackageManager.png)
3. In the **Package Manager Console**, select the project (**Microsoft-Graph-Mail-Console-App**) as the **Default project**.
4. Enter and run **Install-Package Newtonsoft.Json** to add the Json.NET NuGet package.
5. Enter and run **Install-Package Microsoft.IdentityModel.Clients.ActiveDirectory** to add the ADAL (short for Active Directory Authentication Library) NuGet package. 
   ![](https://raw.githubusercontent.com/simonjaeger/OfficeDev-HOL/master/Microsoft-Graph-Mail-Console-App/Images/PackageManagerConsole.png)

#### Exercise 3.1: Add the models ####
We will use a few different models (objects) when calling the Microsoft Graph: These are
- **UserModel:** Represents the current user in the Microsoft Graph. We will use the **Mail** property to use as the recipient of our mail.
- **MailModel:** Contains a single property, **Message**.
- **MessageModel**: This is the message (or mail of you will) itself. In the real world, it has a bunch of different properties - but we will only expose a few in this hands-on lab.
- **BodyModel**: Contains the body content and the format (HTML in this case).
- **ToRecipientModel**: Represents the recipient, contains a single **EmailAddress** property.
- **EmailAddressModel**: Represents the mail address in the Microsoft Graph.

1. Right-click on the project **Microsoft-Graph-Mail-Console-App**  and choose **Add Class...**. Name it **"UserModel"** and click on the **OK** button.
2. In **UserModel.cs**, replace everything with the following code piece.  
    ```csharp
    using System.Collections.Generic;

    namespace Microsoft_Graph_Mail_Console_App
    {
        public class UserModel
        {
            public string Id { get; set; }
            public string UserPrincipalName { get; set; }
            public List<string> BusinessPhones { get; set; }
            public string DisplayName { get; set; }
            public string GivenName { get; set; }
            public object JobTitle { get; set; }
            public string Mail { get; set; }
            public string MobilePhone { get; set; }
            public object OfficeLocation { get; set; }
            public string PreferredLanguage { get; set; }
            public string Surname { get; set; }
        }
    }
    
    ```
2. Right-click on the project **Microsoft-Graph-Mail-Console-App**  and choose **Add Class...**. Name it **"MailModel"** and click on the **OK** button.
3. In **MailModel.cs**, replace everything with the following code piece.  
    ```csharp
    using System.Collections.Generic;

    namespace Microsoft_Graph_Mail_Console_App
    {
        public class MailModel
        {
            public MessageModel Message { get; set; }
        }

        public class MessageModel
        {
            public string Subject { get; set; }
            public BodyModel Body { get; set; }
            public List<ToRecipientModel> ToRecipients { get; set; }
        }

        public class BodyModel
        {
            public string ContentType { get; set; }
            public string Content { get; set; }
        }

        public class ToRecipientModel
        {
            public EmailAddressModel EmailAddress { get; set; }
        }

        public class EmailAddressModel
        {
            public string Address { get; set; }
        }
    }

    ```

#### Exercise 3.2: Add the MailClient class ####
For the next exercises, let's add the following structure for the mail client to better organize where to put our code. 

1. Right-click on the project **Microsoft-Graph-Mail-Console-App**  and choose **Add Class...**. Name it **"MailClient"** and click on the **OK** button.
2. In **MailClient.cs**, replace everything with the following code piece.  
    ```csharp
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

            // The client Id of your native Azure AD application
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
                // Exercise: Get the access token
            }

            // Prepare an HttpClient with the an authorization header (access token)
            public static HttpClient GetHttpClient(string accessToken)
            {
                // Exercise: Get the HttpClient
            }

            // Get the current user using a prepared HttpClient (with an authorization header)
            public static async Task<UserModel> GetUserAsync(HttpClient httpClient)
            {
                // Exercise: Get the user
            }

            // Create the mail structure needed to make a request to the Microsoft Graph
            public static MailModel CreateMail(string subject, string body, params string[] recipients)
            {
                // Exercise: Create the mail
            }

            // Send a mail using a prepared HttpClient (with an authorization header)
            private static async Task<bool> SendMailAsync(HttpClient httpClient, MailModel mail)
            {
                // Exercise: Send the mail
            }

            // Send a mail to the current user
            public static async Task SendMeAsync()
            {
                // Exercise: Assemble everything
            }
        }
    }

    ```
    
#### Exercise 3.3: Configure the MailClient class ####
We need to get back into our Azure AD application configuration page to extract a few properties and put them into the **MailClient.cs** file.

1. In **MailClient.cs**, replace **"[YOUR DOMAIN]"** in the **Domain** property with the domain of your Office 365 Developer Tenant (e.g. contoso.onmicrosoft.com).   
    ```csharp
    // The Office 365 domain (e.g. contoso.microsoft.com)
    public static string Domain
    {
        get { return "simonj.onmicrosoft.com"; }
    }

    ```
2. Head to the configuration page of your Azure AD application. In the **Configure** tab section, locate the **Client ID** property. Copy it.
   ![](https://raw.githubusercontent.com/simonjaeger/OfficeDev-HOL/master/Microsoft-Graph-Mail-Console-App/Images/ClientID.png) 
3. In **MailClient.cs**, replace **"[YOUR CLIENT ID]"** in the **ClientId** property with the client Id of your Azure AD application.   
    ```csharp
    // The client Id of your native Azure AD application
    public static string ClientId
    {
        get { return "deb79365-f124-4719-bc71-f5821cb172dd"; }
    }

    ```
4. Head to the configuration page of your Azure AD application. In the **Configure** tab section, locate the **Redirect URIs** property. Copy it.
   ![](https://raw.githubusercontent.com/simonjaeger/OfficeDev-HOL/master/Microsoft-Graph-Mail-Console-App/Images/RedirectURI.png) 
5. In **MailClient.cs**, replace **"[YOUR REDIRECT URI]"** in the **RedirectUri** property with the redirect URI of your Azure AD application.   
    ```csharp
    // The redirect URI specified in the Azure AD application
    public static Uri RedirectUri
    {
        get { return new Uri("https://Microsoft-Graph-Mail-Console-App/"); }
    }

    ```


#### Exercise 3.4: Finish the MailClient class ####
In order to complete the mail client we need to implement a few methods.
- **GetAccessToken:** Prompt the user with UI (if needed) in order to sign into the application and return an access token (for the Microsoft Graph) from Azure AD.
- **GetHttpClient:** Prepare an **HttpClient** object with the access token attached to the authorization header.
- **GetUserAsync:** Get the current user profile (**UserModel**) using the Microsoft Graph. 
- **CreateMail:** Create the mail object (**MailModel**) which we will send to the Microsoft Graph.
- **SendMailAsync:** Send a mail object (**MailModel**) to the Microsoft Graph. 
- **SendMeAsync:** Assemble the above methods into one single method.

1. In **MailClient.cs**, locate the "Exercise: Get access token" comment and add the following code piece below it.   
    ```csharp
    // Create the authentication context (ADAL)
    var authenticationContext = new AuthenticationContext(Authority);

    // Get the access token
    var authenticationResult = authenticationContext.AcquireToken(GraphResource,
        ClientId, RedirectUri, PromptBehavior.RefreshSession);
    var accessToken = authenticationResult.AccessToken;
    return accessToken;

    ```
2. In **MailClient.cs**, locate the "Exercise: Get HttpClient" comment and add the following code piece below it.   
    ```csharp
    // Create the HTTP client with the access token
    var httpClient = new HttpClient();
    httpClient.DefaultRequestHeaders.Authorization =
        new AuthenticationHeaderValue("Bearer",
        accessToken);
    return httpClient;

    ```   
3. In **MailClient.cs**, locate the "Exercise: Get user" comment and add the following code piece below it.   
    ```csharp
    // Get and deserialize the user
    var userResponse = await httpClient.GetStringAsync(GraphResource + GraphVersion + "/me/");
    var user = JsonConvert.DeserializeObject<UserModel>(userResponse);
    return user;

    ```   
4. In **MailClient.cs**, locate the "Exercise: Create mail" comment and add the following code piece below it.   
    ```csharp
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

    ```   
5. In **MailClient.cs**, locate the "Exercise: Send mail" comment and add the following code piece below it.   
    ```csharp
    // Serialize the mail
    var stringContent = JsonConvert.SerializeObject(mail);

    // Send using POST
    var response = await httpClient.PostAsync(GraphResource + GraphVersion + "/me/microsoft.graph.sendMail",
        new StringContent(stringContent, Encoding.UTF8, "application/json"));
    return response.IsSuccessStatusCode;

    ```      
6. In **MailClient.cs**, locate the "Exercise: Assemble everything" comment and add the following code piece below it.   
    ```csharp
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

    ```  
7. Your **MailClient.cs** file should look like this (except for the Azure AD properties, use your own):   
    ```csharp
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
                get { return "simonj.onmicrosoft.com"; }
            }

            // The authority for authentication; combining the AADInstance
            // and the domain.
            public static string Authority
            {
                get { return string.Format("{0}/{1}/", AADInstance, Domain); }
            }

            // The client Id of your native Azure AD application
            public static string ClientId
            {
                get { return "deb79365-f124-4719-bc71-f5821cb172dd"; }
            }

            // The redirect URI specified in the Azure AD application
            public static Uri RedirectUri
            {
                get { return new Uri("https://Microsoft-Graph-Mail-Console-App/"); }
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
    ```   

#### Exercise 3.5: Finish the Program class ####
In order to trigger the mail client to send ourselves a mail, we need to call it from the **Main** method.

1. In **Program.cs**, mark the **Main** method with the **[STAThread]** attribute.   
    ```csharp
    [STAThread]
    static void Main(string[] args)
    {
       
    }
    
    ```   
2. In **Program.cs**, add the following code piece inside the **Main** method.   
    ```csharp
    MailClient.SendMeAsync().Wait();
    Console.Read();
    
    ```   
3. Your **Program.cs** file should look like this:   
    ```csharp
    using System;

    namespace Microsoft_Graph_Mail_Console_App
    {
        class Program
        {
            [STAThread]
            static void Main(string[] args)
            {
                MailClient.SendMeAsync().Wait();
                Console.Read();
            }
        }
    }

    ```   

> STAThreadAttribute indicates that the COM threading model for the application is single-threaded apartment. This attribute must be present on the entry point of any application that uses Windows Forms; if it is omitted, the Windows components might not work correctly. If the attribute is not present, the application uses the multithreaded apartment model, which is not supported for Windows Forms. <https://msdn.microsoft.com/en-us/library/ms182351(VS.80).aspx>

Launch the project; open the **Debug** menu at the top of Visual Studio 2015 and click on the **Start Debugging** button. You can also click the **Start** button in your toolbar or use the **{F5}** keyboard shortcut.            
   ![](https://raw.githubusercontent.com/simonjaeger/OfficeDev-HOL/master/Excel-Add-in/Images/StartProject.png)
   
At first, you should be asked to sign in using a user in your Office 365 Developer Tenant.
            
   ![](https://raw.githubusercontent.com/simonjaeger/OfficeDev-HOL/master/Microsoft-Graph-Mail-Console-App/Images/SignIn.png)
   
This should generate a mail in your console app and send it to yourself. At this point, go ahead and check your mailbox!
            
   ![](https://raw.githubusercontent.com/simonjaeger/OfficeDev-HOL/master/Microsoft-Graph-Mail-Console-App/Images/NewMail.png)

# Wrap up  #
View the source code files included in this hands-on lab for a final reference of how your code should be structured (if needed). You should now have grasped an understanding of a few possibilities of interacting (and authenticating) with the Microsoft Graph.

# More Resources #
- Discover Office development: <https://msdn.microsoft.com/en-us/office/>
- Learn more about the Microsoft Graph: <http://graph.microsoft.io/en-us/>
- Microsoft Graph: Authentication with Azure AD: <http://simonjaeger.com/microsoft-graph-authentication-with-azure-ad/>
- Microsoft Graph: Authentication with the converged model (preview): <http://simonjaeger.com/microsoft-graph-authentication-with-the-converged-model-preview/>