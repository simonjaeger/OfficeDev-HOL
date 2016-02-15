# Hands-on Lab: Compose Mode Outlook Add-in #

With the new application model for Office comes a brand new way of extending Office with your own functionality - using the tools and dev stacks that we already know and love. 

This hands-on lab demonstrates a few different ways to interact with the Office context.  Setting different properties of a mailbox item (message or appointment) in compose mode. In addition, different styles and components from the Office UI Fabric library is used throughout this Office add-in. 
The objective is to get familiar with some of the possiblities that we have when building Excel add-ins.

The hands-on lab is divided into multiple exercises and should be followed in a chronological order. These are the included exercises:

* [1.1 Create the project](#exercise-11-create-the-project)
* [1.2 Edit the manifest](#exercise-12-edit-the-manifest)
* [1.3 Launch the project](#exercise-13-launch-the-project)
* [2.1 Clean up the project](#exercise-21-clean-up-the-project)
* [2.2 Add Office UI Fabric](#exercise-22-add-office-ui-fabric)
* [2.3 Add the base](#exercise-23-add-the-base-css--html)
* [3.1 Set the item subject](#exercise-31-set-the-item-subject)
* [3.2 Set the item recipients](#exercise-32-set-the-item-recipients)
* [3.3 Set the item body](#exercise-33-set-the-item-body)

Short of time and just want the final sample? Clone this repository (```git clone https://github.com/simonjaeger/OfficeDev-HOL.git```) and open the solution file: **Compose-Mode-Outlook-Add-in\\Source\\Compose-Mode-Outlook-Add-in.sln**.           
    
### Applies to ###
-  Outlook Client
-  Outlook Online (OWA)

### Prerequisites ###
- Visual Studio 2015: <https://www.visualstudio.com/en-us/downloads/download-visual-studio-vs.aspx>
- Office Developer Tools: <https://www.visualstudio.com/en-us/features/office-tools-vs.aspx>
- Office 2013 (Service Pack 1) or Office 2016
- Office 365 Developer Tenant: <http://dev.office.com/devprogram>

### Solution ###
Solution | Author(s)
---------|----------
Compose-Mode-Outlook-Add-in | Simon Jäger (**Microsoft**)

### Version history ###
Version  | Date | Comments
---------| -----| --------
1.0  | February 14th 2016 | Initial release

### Disclaimer ###
**THIS CODE IS PROVIDED *AS IS* WITHOUT WARRANTY OF ANY KIND, EITHER EXPRESS OR IMPLIED, INCLUDING ANY IMPLIED WARRANTIES OF FITNESS FOR A PARTICULAR PURPOSE, MERCHANTABILITY, OR NON-INFRINGEMENT.**

----------

# Exercises #
#### Exercise 1.1: Create the project ####
The first thing that we need to do is to create the project itself. Make sure that you have installed all of the required prerequisites before launching Visual Studio 2015. 

1. Click **File**, **New** and finally the **Project** button.
2. In **Templates**, select **Visual C#**, **Office/SharePoint** and then **Office Add-ins**. This will list the Office add-in templates, choose **Office Add-in**. 
3. Name your project **"Compose-Mode-Outlook-Add-in"** and click the **OK** button to continue. 
   ![](https://raw.githubusercontent.com/simonjaeger/OfficeDev-HOL/master/Read-Mode-Outlook-Add-in/Images/NewProject.png)
4. Next up Visual Studio 2015 will need a bit more information about what you are going to create - in order to set up the required files. Your next step is to decide which type of Office add-in that you want to create. Depending on what you pick, your Office add-in will run in different Office applications and contexts. 
   
   For this hands-on lab, we will create a mail add-in - this means that our Office add-in will run in in Outlook as a view beside the Office context (e.g. a message or appointment). Select **Mail** and click on **Next**. 
   ![](https://raw.githubusercontent.com/simonjaeger/OfficeDev-HOL/master/Read-Mode-Outlook-Add-in/Images/MailAddinType.png)
5. Finally we need to choose the supported modes for our mail add-in. This means that we are defining the contexts that our mail add-in can run within; read, compose or both. If you choose **Read form**, the mail add-in will be able to run when a user is viewing a mailbox item. In **Compose form**, the mail add-in can run when a user is creating or editing a mailbox item. 
   
   In our case, select **Compose form** for both **Email message** and **Appointment**. Deselect everything else to create a compose mode mail add-in. Click **Finish** to complete the wizard.
   ![](https://raw.githubusercontent.com/simonjaeger/OfficeDev-HOL/master/Compose-Mode-Outlook-Add-in/Images/ComposeMailAddin.png)
6. Using the information you specified in the wizard, Visual Studio 2015 will configure your project. Have a look in the **Solution Explorer** and find your two new projects in the **Compose-Mode-Outlook-Add-in** solution. 
   
   **Compose-Mode-Outlook-Add-in:** This is your manifest project, containing the XML manifest file. This is basically a representation of the information you just specified while creating your Office add-in project. 
   
   **Compose-Mode-Outlook-Add-inWeb:** This is your web project for the Office add-in. This contains all of the different source files that makes up your Outlook add-in. We will make quite a few adjustments to this structure as we continue.
   ![](https://raw.githubusercontent.com/simonjaeger/OfficeDev-HOL/master/Compose-Mode-Outlook-Add-in/Images/SolutionExplorer.png)
   
   You've now created the basic structure for a mail add-in running in Outlook. 

#### Exercise 1.2: Edit the manifest ####
We need to make sure that we understand the manifest file. This file is essential for your add-in; it tells Office where everything is hosted (locally throughout this hands-on lab) and where it can be launched. So let's open it and edit the manifest file.

1. In the manifest project **Compose-Mode-Outlook-Add-in**, double-click the **Compose-Mode-Outlook-Add-inManifest** file. This will open the manifest editor.      
   ![](https://raw.githubusercontent.com/simonjaeger/OfficeDev-HOL/master/Compose-Mode-Outlook-Add-in/Images/OutlookAddinManifest.png)
2. In the **General** tab section, find and edit the **Display name** and **Provider name** to anything you'd like.      
   ![](https://raw.githubusercontent.com/simonjaeger/OfficeDev-HOL/master/Compose-Mode-Outlook-Add-in/Images/GeneralAddinManifest.png)
3. In the **Compose Form** tab section, find the **Activation** part. This is what determines the potential activation of your mail add-in. By default, both **Email messages** and **Appointments** should be checked. 
4. Pay attention down below to the **Source location** property. This points to a specific file in your web project (**Compose-Mode-Outlook-Add-inWeb**). When launching your mail add-in, this page will be the first thing that gets loaded and displayed.       
   ![](https://raw.githubusercontent.com/simonjaeger/OfficeDev-HOL/master/Compose-Mode-Outlook-Add-in/Images/ComposeFormAddinManifest.png)

#### Exercise 1.3: Launch the project ####
Before we launch our mail add-in we should validate that our start actions are proper.

1. Select the manifest project; **Compose-Mode-Outlook-Add-in** in the **Solution Explorer**.                                     
   ![](https://raw.githubusercontent.com/simonjaeger/OfficeDev-HOL/master/Compose-Mode-Outlook-Add-in/Images/SelectManifestProject.png)
2. In the **Properties** window, set **Start Action** to Office Desktop Client. 
4. Set **Web Project** to your web project; **Compose-Mode-Outlook-Add-inWeb**.
   ![](https://raw.githubusercontent.com/simonjaeger/OfficeDev-HOL/master/Compose-Mode-Outlook-Add-in/Images/StartActions.png)
5. To launch the project, open on the **Debug** menu at the top of Visual Studio 2015 and click on the **Start Debugging** button. You can also click the **Start** button in your toolbar or use the **{F5}** keyboard shortcut.            
   ![](https://raw.githubusercontent.com/simonjaeger/OfficeDev-HOL/master/Excel-Add-in/Images/StartProject.png)
6. When launching your mail add-in for the first time, Visual Studoo 2015 needs to install the manifest file. This is where you should use your Office 365 Developer Tenant (if you haven't signed up for one yet, get yours for free at <http://dev.office.com/devprogram>). Enter the credentials of a user (**[username]@[your domain].onmicrosoft.com**) belonging to your Office 365 Developer Tenant and click on the **Connect** button.            
   ![](https://raw.githubusercontent.com/simonjaeger/OfficeDev-HOL/master/Read-Mode-Outlook-Add-in/Images/ConnectToExchange.png)
7. Once Outlook has launched, you'll notice that your mail add-in doesn't start right away. We need to start it manually. Create a new message and click on the **Office Add-ins** button.             
   ![](https://raw.githubusercontent.com/simonjaeger/OfficeDev-HOL/master/Compose-Mode-Outlook-Add-in/Images/AddAddins.png)
8. Select **Compose-Mode-Outlook-Add-in** and click **Start**.              
   ![](https://raw.githubusercontent.com/simonjaeger/OfficeDev-HOL/master/Compose-Mode-Outlook-Add-in/Images/AddAddin.png)
9. Once your mail add-in has launched, you can explore the functionality that comes right out of the box with the Visual Studio 2015 template.            
   ![](https://raw.githubusercontent.com/simonjaeger/OfficeDev-HOL/master/Compose-Mode-Outlook-Add-in/Images/LaunchedComposeMailAddin.png)
10. Finally, stop debugging by opening the **Debug** menu at the top of Visual Studio 2015 and click on the **Stop Debugging** button. You can also click the **Stop** button in your toolbar.            
   ![](https://raw.githubusercontent.com/simonjaeger/OfficeDev-HOL/master/Word-Add-in/Images/StopDebugging.png)

#### Exercise 2.1: Clean up the project ####
While the default styling that comes along with the Visual Studio 2015 template for Office add-ins does its job - leveraging the features of the Office UI Fabric can be fantastic. It's a UI toolkit made specifically for building Office and Office 365 experiences, so it will certainly help us out here.

The Office UI Fabric library comes with everything from styling, components to animations. The majority of the library can be references via a CDN. The heavier parts needs to be downloaded and added to the project itself. We will go through both of these approaches. 

Our first task is to clean up the project, and remove the default styling and setup.

1. Remove the **Content** and **Images** folders from the web project. You can do this by right-clicking these folders in the **Solution Explorer** and choosing the **Delete** option.                                    
    ![](https://raw.githubusercontent.com/simonjaeger/OfficeDev-HOL/master/Word-Add-in/Images/DeleteFolders.png)
2. In your **Solution Explorer**, find the **Home.html** file - which is the startup page for your mail add-in. **Remove** everything inside the **body** tags. 
3. In **Home.html** remove the CSS reference to **"../../Content/Office.css"** - we have removed this file and will be using Office UI Fabric instead. This should leave you with this:
    ```html
    <!DOCTYPE html>
    <html>
    <head>
        <meta charset="UTF-8" />
        <meta http-equiv="X-UA-Compatible" content="IE=Edge" />
        <title></title>
        <script src="../../Scripts/jquery-1.9.1.js" type="text/javascript"></script>

        <script src="https://appsforoffice.microsoft.com/lib/1/hosted/office.js" type="text/javascript"></script>

        <link href="../App.css" rel="stylesheet" type="text/css" />
        <script src="../App.js" type="text/javascript"></script>

        <link href="Home.css" rel="stylesheet" type="text/css" />
        <script src="Home.js" type="text/javascript"></script>
    </head>
    <body>
    
    </body>
    </html>

    ```
4. In **App.js**, **remove** the **initialize()** function defined on the **app** object, as this will not be used:            
    ```js
    var app = (function () {
        "use strict";
    
        var app = {};
        return app;
    })();
    ```
5. In **Home.js**, remove the **setSubject()**, **getSubject()** and **addToRecipients()** functions (and the corresponding click event handler registrations in the **document.ready** function). 
6. In **Home.js**, remove the call to **app.initialize()**. We are remaking the structure of the mail add-in, these will no longer be used. You should end up with this:            
    ```js
    (function () {
        "use strict";
    
        // The initialize function must be run each time a new page is loaded
        Office.initialize = function (reason) {
            $(document).ready(function () {
            });
        };
    })();
    ```
6. In **App.css**, **remove** everything, leaving you with an empty file.
7. In **Home.css**, **remove** everything, leaving you with an empty file.

#### Exercise 2.2: Add Office UI Fabric ####
1. In **Home.html**, add two CSS references to the CDN for Office UI Fabric inside the **head** tags. Add them before the CSS reference to **"../App.css"**.
    ```html
    <link rel="stylesheet" href="https://appsforoffice.microsoft.com/fabric/1.0/fabric.min.css">
    <link rel="stylesheet" href="https://appsforoffice.microsoft.com/fabric/1.0/fabric.components.min.css">
    
    ```
    
2. Your **Home.html** file should now look like this: 
    ```html
    <!DOCTYPE html>
    <html>
    <head>
        <meta charset="UTF-8" />
        <meta http-equiv="X-UA-Compatible" content="IE=Edge" />
        <title></title>
        <script src="../../Scripts/jquery-1.9.1.js" type="text/javascript"></script>

        <script src="https://appsforoffice.microsoft.com/lib/1/hosted/office.js" type="text/javascript"></script>
    
        <link rel="stylesheet" href="https://appsforoffice.microsoft.com/fabric/1.0/fabric.min.css">
        <link rel="stylesheet" href="https://appsforoffice.microsoft.com/fabric/1.0/fabric.components.min.css">
        <script src="../../Scripts/Jquery.Dropdown.js" type="text/javascript"></script>

        <link href="../App.css" rel="stylesheet" type="text/css" />
        <script src="../App.js" type="text/javascript"></script>
    
        <link href="Home.css" rel="stylesheet" type="text/css" />
        <script src="Home.js" type="text/javascript"></script>
    </head>
    <body>
    
    </body>
    </html>
    ``` 

#### Exercise 2.3: Add the base (CSS + HTML) ####
1. In **App.css**, add the following basic CSS (this should be entire file). We will do much of the styling through already defined classes in the Office UI Fabric. But some basic layouting will do us great!
    ```css
    #header {
        position: absolute;
        top: 0px;
        left: 0px;
        height: 80px;
        width: 100%;
        overflow: hidden;
    }

    #header h2 {
        position: relative;
        margin-left: 22px;
        margin-top: 20px;
        text-wrap: none;
        white-space: nowrap;
        line-height: 22px;
    }

    #content {
        position: absolute;
        top: 80px;
        padding-left: 15px;
        padding-right: 15px;
    }

    .section-title {
        margin-bottom: -5px;
    }

    .section {
        margin-bottom: 10px;
    }
    ``` 
2. In **Home.html**, add the following chunk of HTML inside the **body** tags. This will set the stage for the next array of exercises to come. 
    ```html
    <!-- Header -->
    <div id="header" class="ms-bgColor-themePrimary">
        <h2 class="ms-font-xl ms-fontWeight-semibold ms-fontColor-white">
            HOL: Outlook Add-in
            <br />
            (Compose Mode)
        </h2>
    </div>
    <div id="content">
        <!-- Introduction -->
        <p class="ms-font-m ms-fontColor-neutralSecondary">
            This sample demonstrates a few different ways to interact with the Office context.
            Setting different properties of a mailbox item (message or appointment) in compose mode.
        </p>

        <!-- Exercise Section: Set subject -->
        <p class="ms-font-l ms-fontWeight-semibold section-title">Exercise: Set subject</p>
        <p class="ms-font-m ms-fontColor-neutralSecondary">
            Set the item subject by clicking the button down below.
        </p>

        <!-- Exercise: Set subject -->
        <div class="section">
        </div>

        <!-- Exercise Section: Set recipients -->
        <p class="ms-font-l ms-fontWeight-semibold section-title">Exercise: Set recipients</p>
        <p class="ms-font-m ms-fontColor-neutralSecondary">
            Set the item recipients by clicking the button down below.
        </p>

        <!-- Exercise: Set recipients -->
        <div class="section">
        </div>

        <!-- Exercise Section: Set body -->
        <p class="ms-font-l ms-fontWeight-semibold section-title">Exercise: Set body</p>
        <p class="ms-font-m ms-fontColor-neutralSecondary">
            Set the item body by clicking the button down below.
            This requires a minimum mailbox requirement set version of 1.3.
        </p>

        <!-- Exercise: Set body -->
        <div class="section">
        </div>

        <!-- Office UI Fabric -->
        <p class="ms-font-l ms-fontWeight-semibold section-title">Office UI Fabric</p>
        <p class="ms-font-m ms-fontColor-neutralSecondary">
            Different styles and components from the Office UI Fabric library is used throughout this Office add-in.
        </p>
        <p class="ms-font-m ms-fontColor-neutralSecondary">
            Learn more about Office UI Fabric at: <a class="ms-Link" href="http://dev.office.com/fabric/" target="_blank">http://dev.office.com/fabric/</a>
        </p>
    </div>
    ``` 
3. Launch your mail add-in to display the new styling. We will add more interactive components in the different sections (inside the recently added HTML piece).            
   ![](https://raw.githubusercontent.com/simonjaeger/OfficeDev-HOL/master/Compose-Mode-Outlook-Add-in/Images/LaunchedComposeMailAddin2.png)

#### Exercise 3.1: Set the item subject ####
1. In **Home.html**, locate the "Exercise: Set subject" section (commented) and add the following HTML piece inside the **div** (section) tags. This is an Office UI Fabric styled button. 
    ```html
    <button id="set-subject" class="ms-Button">
        <span class="ms-Button-label">Set subject</span>
    </button>
    
    ```
2. In **Home.js**, add an event handler (below the initialization of the Office UI Fabric components, in the **ready** function) for the click event of the button:
    ```js
    // Add event handlers
    $('#set-subject').click(setSubject);
    
    ```
3. In **Home.js**, add the following functions to set the item subject (below the **Office.initialize** function):
    ```js
    // Set the item subject
    function setSubject() {
        var _item = Office.context.mailbox.item;
        var subject = _item.subject;
        subject.setAsync('Hello #Office365Dev', onDataSet)
    }

    // Callback function for the asynchronous write functions
    function onDataSet(asyncResult) {
        if (asyncResult.status !== Office.AsyncResultStatus.Succeeded) {
            // TODO: Handle error
        }
    }
    ```
4. Launch your mail add-in and test your work by clicking the **Set subject** button. When the button is clicked, the function will be executed; setting the item subject in the mailbox item.

#### Exercise 3.2: Set the item recipients ####
1. In **Home.html**, locate the "Exercise: Set recipients" section (commented) and add the following HTML piece inside the **div** (section) tags. This is an Office UI Fabric styled button. 
    ```html
    <button id="set-recipients" class="ms-Button">
        <span class="ms-Button-label">Set recipients</span>
    </button>
    
    ```
2. In **Home.js**, add an event handler (below the initialization of the Office UI Fabric components, in the **ready** function) for the click event of the button:
    ```js
    $('#set-recipients').click(setRecipients);
    
    ```
3. In **Home.js**, add the following function to set the item recipients:
    ```js
    // Set the item recipients
    function setRecipients() {
        var _item = Office.context.mailbox.item;
        var _userProfile = Office.context.mailbox.userProfile;
        var user = {
            displayName: _userProfile.displayName,
            emailAddress: _userProfile.emailAddress
        };

        // Check if the item is a message or appointment
        // in order to determine which properties to set
        if (_item.itemType === Office.MailboxEnums.ItemType.Message) {
            _item.to.setAsync([user], onDataSet);
            _item.cc.setAsync([user], onDataSet);
            _item.bcc.setAsync([user], onDataSet);
        }
        else if (_item.itemType === Office.MailboxEnums.ItemType.Appointment) {
            _item.requiredAttendees.setAsync([user], onDataSet);
        }
    }
    ```
4. Launch your mail add-in and test your work by clicking the **Set recipients** button. When the button is clicked, the function will be executed; setting the item recipients of the mailbox item.



#### Exercise 3.3: Set the item body ####
1. In **Home.html**, locate the "Exercise: Set body" section (commented) and add the following HTML piece inside the **div** (section) tags. This is an Office UI Fabric styled button. 
    ```html
    <button id="set-body" class="ms-Button ms-Button--primary">
        <span class="ms-Button-label">Set body</span>
    </button>
    
    ```
2. In **Home.js**, add an event handler (below the initialization of the Office UI Fabric components, in the **ready** function) for the click event of the button:
    ```js
    $('#set-body').click(setBody);
    
    ```
3. In **Home.js**, add the following function to set the item body (below the **Office.initialize** function):
    ```js
    // Set the item body
    function setBody() {
        var _item = Office.context.mailbox.item;
        var body = _item.body;
        var text = 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod' +
            'tempor incididunt ut labore et dolore magna aliqua. Ut enim ad minim veniam, ' +
            'quis nostrud exercitation ullamco laboris nisi ut aliquip ex ea commodo consequat. ' +
            'Duis aute irure dolor in reprehenderit in voluptate velit esse cillum dolore ' +
            'eu fugiat nulla pariatur. Excepteur sint occaecat cupidatat non proident, sunt ' +
            'in culpa qui officia deserunt mollit anim id est laborum.';
        body.setAsync(text, onDataSet)
    }
    
    ```
4. Setting the item body is an asynchronous function that requires a minimum mailbox requirement set version of 1.3. There are different ways of ensuring that your user has at least version 1.3, a good way is to set it in the manifest.       
         
   In the manifest project **Compose-Mode-Outlook-Add-in**, double-click the **Compose-Mode-Outlook-Add-inManifest** file. This will open the manifest editor.      
   ![](https://raw.githubusercontent.com/simonjaeger/OfficeDev-HOL/master/Compose-Mode-Outlook-Add-in/Images/OutlookAddinManifest.png)
5. In the **General** tab section, find the **Mailbox requirement set** property and set it to **1.3**.       
   ![](https://raw.githubusercontent.com/simonjaeger/OfficeDev-HOL/master/Read-Mode-Outlook-Add-in/Images/RequirementSet.png)
4. Launch your mail add-in and test your work by clicking the **Set body** button. When the button is clicked, the function will be executed; setting the item body.            
   ![](https://raw.githubusercontent.com/simonjaeger/OfficeDev-HOL/master/Compose-Mode-Outlook-Add-in/Images/LaunchedComposeMailAddin3.png)




# Wrap up  #
View the source code files included in this hands-on lab for a final reference of how your code should be structured (if needed). You should now have grasped an understanding of a few possibilities of interacting with the Office context (a mailbox item in this case). In addition, you have also seen some of the styles and components included in the Office UI Fabric.

# More Resources #
- Discover Office development: <https://msdn.microsoft.com/en-us/office/>
- Learn more about Office UI Fabric: <http://dev.office.com/fabric/>
- Developing Outlook add-ins – where to integrate and what you can do: <http://simonjaeger.com/developing-outlook-add-ins-where-to-integrate-and-what-you-can-do/>