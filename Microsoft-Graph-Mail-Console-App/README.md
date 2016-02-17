# Hands-on Lab: Microsoft Graph Mail Console App #

If you haven’t heard, there is an easy way to call a great amount of Microsoft APIs – using one single endpoint. This endpoint, so called the Microsoft Graph (https://graph.microsoft.io/) lets you access everything from data, to intelligence and insights powered by the Microsoft cloud.

No longer will you need to keep track of different endpoints and separate tokens in your solutions – how great is that? This post is an introductory part of getting started with the Microsoft Graph. For changes in the Microsoft Graph, head to: https://graph.microsoft.io/changelog

This hands-on lab demonstrates how to configure Azure AD to acquire an access token, used when calling the Microsoft Graph. With the Microsoft Graph, we will send ourselves a mail from the console application.
The objective is to get familiar with the authentication piece (which is the complex part), but also to display some of the possibilities of the Microsoft Graph.

The hands-on lab is divided into multiple exercises and should be followed in a chronological order. These are the included exercises:

* [1.1 Create the project](#exercise-11-create-the-project)
* [1.2 Edit the manifest](#exercise-12-edit-the-manifest)
* [1.3 Launch the project](#exercise-13-launch-the-project)
* [2.1 Clean up the project](#exercise-21-clean-up-the-project)
* [2.2 Add Office UI Fabric](#exercise-22-add-office-ui-fabric)
* [2.3 Add the base](#exercise-23-add-the-base-css--html)
* [3.1 Add the input fields](#exercise-31-add-the-input-fields)
* 

Short of time and just want the final sample? Clone this repository (```git clone https://github.com/simonjaeger/OfficeDev-HOL.git```) and open the solution file: **Single-Sign-On-Outlook-Add-in\\Source\\Single-Sign-On-Outlook-Add-in.sln**.           
    
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

#### Exercise 2.2: Add the models ####
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

#### Exercise 2.3: Add the client ####

#### Exercise 2.4: Finish the Program class ####

# Wrap up  #
View the source code files included in this hands-on lab for a final reference of how your code should be structured (if needed). You should now have grasped an understanding of a few possibilities of interacting (and authenticating) with the Microsoft Graph.

# More Resources #
- Discover Office development: <https://msdn.microsoft.com/en-us/office/>
- Learn more about the Microsoft Graph: <http://graph.microsoft.io/en-us/>
- Microsoft Graph: Authentication with Azure AD: <http://simonjaeger.com/microsoft-graph-authentication-with-azure-ad/>
- Microsoft Graph: Authentication with the converged model (preview): <http://simonjaeger.com/microsoft-graph-authentication-with-the-converged-model-preview/>