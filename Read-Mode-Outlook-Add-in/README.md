# Hands-on Lab: Read Mode Outlook Add-in #

With the new application model for Office comes a brand new way of extending Office with your own functionality - using the tools and dev stacks that we already know and love. 

This hands-on lab demonstrates a few different ways to interact with the Office context.  Accessing different types of data for a mailbox item (message or appointment) in read mode. In addition, different styles and components from the Office UI Fabric library is used throughout this Office add-in. 
The objective is to get familiar with some of the possiblities that we have when building Excel add-ins.

The hands-on lab is divided into multiple exercises and should be followed in a chronological order. These are the included exercises:

* [1.1 Create the project](#exercise-11-create-the-project)
* [1.2 Edit the manifest](#exercise-12-edit-the-manifest)
* [1.3 Launch the project](#exercise-13-launch-the-project)
* [2.1 Clean up the project](#exercise-21-clean-up-the-project)
* [2.2 Add Office UI Fabric](#exercise-22-add-office-ui-fabric)
* [2.3 Add the base](#exercise-23-add-the-base-css--html)
* [3.1 Add a dialog](#exercise-31-add-a-dialog)
* [4.1 Get the item subject](#exercise-41-get-the-item-subject)
* [4.2 Get the item sender](#exercise-42-get-the-item-sender)
* [4.3 Get the item body](#exercise-43-get-the-item-body)

Short of time and just want the final sample? Clone this repository (```git clone https://github.com/simonjaeger/OfficeDev-HOL.git```) and open the solution file: **Read-Mode-Outlook-Add-in\\Source\\Read-Mode-Outlook-Add-in.sln**.           
    
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
Read-Mode-Outlook-Add-in | Simon Jäger (**Microsoft**)

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
3. Name your project **"Read-Mode-Outlook-Add-in"** and click the **OK** button to continue. 
   ![](https://raw.githubusercontent.com/simonjaeger/OfficeDev-HOL/master/Read-Mode-Outlook-Add-in/Images/NewProject.png)
4. Next up Visual Studio 2015 will need a bit more information about what you are going to create - in order to set up the required files. Your next step is to decide which type of Office add-in that you want to create. Depending on what you pick, your Office add-in will run in different Office applications and contexts. 
   
   For this hands-on lab, we will create a mail add-in - this means that our Office add-in will run in in Outlook as a view beside the Office context (e.g. a message or appointment). Select **Mail** and click on **Next**. 
   ![](https://raw.githubusercontent.com/simonjaeger/OfficeDev-HOL/master/Read-Mode-Outlook-Add-in/Images/MailAddinType.png)
5. Finally we need to choose the supported modes for our mail add-in. This means that we are defining the contexts that our mail add-in can run within; read, compose or both. If you choose **Read form**, the mail add-in will be able to run when a user is viewing a mailbox item. In **Compose form**, the mail add-in can run when a user is creating or editing a mailbox item. 
   
   In our case, select **Read form** for both **Email message** and **Appointment**. Deselect everything else to create a read mode mail add-in. Click **Finish** to complete the wizard.
   ![](https://raw.githubusercontent.com/simonjaeger/OfficeDev-HOL/master/Read-Mode-Outlook-Add-in/Images/ReadMailAddin.png)
6. Using the information you specified in the wizard, Visual Studio 2015 will configure your project. Have a look in the **Solution Explorer** and find your two new projects in the **Read-Mode-Outlook-Add-in** solution. 
   
   **Read-Mode-Outlook-Add-in:** This is your manifest project, containing the XML manifest file. This is basically a representation of the information you just specified while creating your Office add-in project. 
   
   **Read-Mode-Outlook-Add-inWeb:** This is your web project for the Office add-in. This contains all of the different source files that makes up your Outlook add-in. We will make quite a few adjustments to this structure as we continue.
   ![](https://raw.githubusercontent.com/simonjaeger/OfficeDev-HOL/master/Read-Mode-Outlook-Add-in/Images/SolutionExplorer.png)
   
   You've now created the basic structure for a mail add-in running in Outlook. 

#### Exercise 1.2: Edit the manifest ####
We need to make sure that we understand the manifest file. This file is essential for your add-in; it tells Office where everything is hosted (locally throughout this hands-on lab) and where it can be launched. So let's open it and edit the manifest file.

1. In the manifest project **Read-Mode-Outlook-Add-in**, double-click the **Read-Mode-Outlook-Add-inManifest** file. This will open the manifest editor.      
   ![](https://raw.githubusercontent.com/simonjaeger/OfficeDev-HOL/master/Read-Mode-Outlook-Add-in/Images/OutlookAddinManifest.png)
2. In the **General** tab section, find and edit the **Display name** and **Provider name** to anything you'd like.      
   ![](https://raw.githubusercontent.com/simonjaeger/OfficeDev-HOL/master/Read-Mode-Outlook-Add-in/Images/GeneralAddinManifest.png)
3. In the **Read Form** tab section, find the **Activation** part. This is what determines the rules for potential activation of your mail add-in. By default, both **Item is a message** and **Item is an appointment** should be included. 
4. Scroll down and pay attention to the **Source location** property. This points to a specific file in your web project (**Read-Mode-Outlook-Add-inWeb**). When launching your mail add-in, this page will be the first thing that gets loaded and displayed.       
   ![](https://raw.githubusercontent.com/simonjaeger/OfficeDev-HOL/master/Read-Mode-Outlook-Add-in/Images/ReadFormAddinManifest.png)

#### Exercise 1.3: Launch the project ####
Before we launch our mail add-in we should validate that our start actions are proper.


1. Select the manifest project; **Read-Mode-Outlook-Add-in** in the **Solution Explorer**.                                     
   ![](https://raw.githubusercontent.com/simonjaeger/OfficeDev-HOL/master/Read-Mode-Outlook-Add-in/Images/SelectManifestProject.png)
2. In the **Properties** window, set **Start Action** to Office Desktop Client. 
4. Set **Web Project** to your web project; **Read-Mode-Outlook-Add-inWeb**.
   ![](https://raw.githubusercontent.com/simonjaeger/OfficeDev-HOL/master/Read-Mode-Outlook-Add-in/Images/StartActions.png)
5. To launch the project, open on the **Debug** menu at the top of Visual Studio 2015 and click on the **Start Debugging** button. You can also click the **Start** button in your toolbar or use the **{F5}** keyboard shortcut.            
   ![](https://raw.githubusercontent.com/simonjaeger/OfficeDev-HOL/master/Excel-Add-in/Images/StartProject.png)
6. When launching your mail add-in for the first time, Visual Studoo 2015 needs to install the manifest file. This is where you should use your Office 365 Developer Tenant (if you haven't signed up for one yet, get yours for free at <http://dev.office.com/devprogram>). Enter the credentials of a user (**[username]@[your domain].onmicrosoft.com**) belonging to your Office 365 Developer Tenant and click on the **Connect** button.            
   ![](https://raw.githubusercontent.com/simonjaeger/OfficeDev-HOL/master/Read-Mode-Outlook-Add-in/Images/ConnectToExchange.png)
7. Once Outlook has launched, you'll notice that your mail add-in doesn't start right away. We need to start it manually. Select a message in your mailbox (send yourself one if needed) and click on the **Read-Mode-Outlook-Add-in** above it. Once your mail add-in has launched, you can explore the functionality that comes right out of the box with the Visual Studio 2015 template.            
   ![](https://raw.githubusercontent.com/simonjaeger/OfficeDev-HOL/master/Read-Mode-Outlook-Add-in/Images/LaunchedReadMailAddin.png)
8. Finally, stop debugging by opening the **Debug** menu at the top of Visual Studio 2015 and click on the **Stop Debugging** button. You can also click the **Stop** button in your toolbar.            
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
5. In **Home.js**, **remove** the **displayItemDetails()** function and the call to **app.initialize()**. We are remaking the structure of the mail add-in, these will no longer be used. You should end up with this:            
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
        margin-top: 30px;
        text-wrap: none;
        white-space: nowrap;
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
        <h2 class="ms-font-xxl ms-fontWeight-semibold ms-fontColor-white">HOL: Outlook Add-in (Read Mode)</h2>
    </div>
    <div id="content">
        <!-- Introduction -->
        <p class="ms-font-m ms-fontColor-neutralSecondary">
            This sample demonstrates a few different ways to interact with the Office context.
            Accessing different types of data for a mailbox item (message or appointment) in read mode.
        </p>

        <div class="ms-Grid">
            <div class="ms-Grid-row">
                <!-- Exercise Section: Get subject -->
                <div class="ms-Grid-col ms-u-sm6 ms-u-md6" style="padding-left: 0px;">
                    <p class="ms-font-l ms-fontWeight-semibold section-title">Exercise: Get subject</p>
                    <p class="ms-font-m ms-fontColor-neutralSecondary">
                        Get the item subject and display it by clicking the button down below.
                    </p>

                    <!-- Exercise: Get subject -->

                </div>

                <!-- Exercise Section: Get sender -->
                <div class="ms-Grid-col ms-u-sm6 ms-u-md6">
                    <p class="ms-font-l ms-fontWeight-semibold section-title">Exercise: Get sender</p>
                    <p class="ms-font-m ms-fontColor-neutralSecondary">
                        Get the item sender and display it by clicking the button down below.
                    </p>

                    <!-- Exercise: Get sender -->

                </div>
            </div>

            <div class="ms-Grid-row">
                <!-- Exercise Section: Get body -->
                <div class="ms-Grid-col ms-u-sm6 ms-u-md6" style="padding-left: 0px;">
                    <p class="ms-font-l ms-fontWeight-semibold section-title">Exercise: Get body</p>
                    <p class="ms-font-m ms-fontColor-neutralSecondary">
                        Get the item body by clicking the button down below.
                        This requires a minimum mailbox requirement set version of 1.3.
                    </p>

                    <!-- Exercise: Get body -->

                </div>

                <!-- Office UI Fabric -->
                <div class="ms-Grid-col ms-u-sm6 ms-u-md6">
                    <p class="ms-font-l ms-fontWeight-semibold section-title">Office UI Fabric</p>
                    <p class="ms-font-m ms-fontColor-neutralSecondary">
                        Different styles and components from the Office UI Fabric library is used throughout this Office add-in.
                    </p>
                    <p class="ms-font-m ms-fontColor-neutralSecondary">
                        Learn more about Office UI Fabric at: <a class="ms-Link" href="http://dev.office.com/fabric/" target="_blank">http://dev.office.com/fabric/</a>
                    </p>
                </div>
            </div>
        </div>

        <!-- Exercise: Dialog -->

    </div>
    ``` 
3. Launch your mail add-in to display the new styling. We will add more interactive components in the different sections (inside the recently added HTML piece).            
   ![](https://raw.githubusercontent.com/simonjaeger/OfficeDev-HOL/master/Read-Mode-Outlook-Add-in/Images/LaunchedReadMailAddin2.png)

#### Exercise 3.1: Add a dialog ####
As we are going to extract property values from the mailbox item, let's have a more sophisticated approach than the **JavaScript Console**. We can use an Office UI Fabric dialog to display this data in a more user-friendly way.  

1. In **Home.html**, locate the "Exercise: Data dialog" section (commented) and add the following HTML piece right after. This is an Office UI Fabric styled dialog. 
    ```html
    <div id="data-dialog" class="ms-Dialog ms-Dialog--lgHeader">
        <div class="ms-Overlay"></div>
        <div class="ms-Dialog-main">
            <button class="ms-Dialog-button ms-Dialog-button--close">
                <i class="ms-Icon ms-Icon--x"></i>
            </button>
            <div class="ms-Dialog-header">
                <p id="data-dialog-title" class="ms-Dialog-title">
                    <!-- Dialog title -->
                </p>
            </div>
            <div class="ms-Dialog-inner">
                <div class="ms-Dialog-content">
                    <p id="data-dialog-text" class="ms-Dialog-subText" style="max-height: 95px; overflow-y: hidden;">
                        <!-- Dialog text -->
                    </p>
                </div>
                <div class="ms-Dialog-actions">
                    <div class="ms-Dialog-actionsRight">
                        <button id="data-dialog-more" class="ms-Dialog-action ms-Button ms-Button--primary">
                            <span class="ms-Button-label">OK</span>
                        </button>
                        <button id="data-dialog-got-it" class="ms-Dialog-action ms-Button">
                            <span class="ms-Button-label">Got it</span>
                        </button>
                    </div>
                </div>
            </div>
        </div>
    </div>
    
    ```
2. Launch your mail add-in and test your work. You should find a dialog covering up most of your display area.            
   ![](https://raw.githubusercontent.com/simonjaeger/OfficeDev-HOL/master/Read-Mode-Outlook-Add-in/Images/LaunchedReadMailAddin3.png)
3. In **Home.js**, add two event handlers (below the initialization of the Office UI Fabric components, in the **ready** function) for the dialog. This will allow us to close it.
    ```js
    // Add event handlers
    $('#data-dialog-more').click(hideDataDialog);
    $('#data-dialog-got-it').click(hideDataDialog);
    
    ```
4. In **Home.js**, add the following functions to show and close the dialog:
    ```js
    // Show the data dialog
    function showDataDialog(title, text) {
        // Set the dialog title and text
        $('#data-dialog-title').text(title);
        $('#data-dialog-text').text(text);
        $('#data-dialog').show();
    }

    // Hide the data dialog
    function hideDataDialog() {
        $('#data-dialog').hide();
    }
    ```
5. Launch your mail add-in and test your work. You should be able to close the dialog using any of the two buttons.
6. In **Home.css**, add the following CSS piece to hide the dialog when your mail add-in has launched.
    ```css
    #data-dialog {
        display: none;
    }
    
    ```

#### Exercise 4.1: Get the item subject ####
1. In **Home.html**, locate the "Exercise: Get subject" section (commented) and add the following HTML piece inside the **div** (section) tags. This is an Office UI Fabric styled button. 
    ```html
    <button id="get-subject" class="ms-Button">
        <span class="ms-Button-label">Get subject</span>
    </button>
    
    ```
2. In **Home.js**, add an event handler (below the initialization of the Office UI Fabric components, in the **ready** function) for the click event of the button:
    ```js
    $('#get-subject').click(getSubject);
    
    ```
3. In **Home.js**, add the following function to get the item subject (below the **Office.initialize** function):
    ```js
    // Get the item subject and display it
    function getSubject() {
        var _item = Office.context.mailbox.item;
        var subject = _item.subject;

        // Show data
        showDataDialog('Subject', subject);
    }
    ```
4. Launch your mail add-in and test your work by clicking the **Get subject** button. When the button is clicked, the function will be executed; displaying the item subject using the data dialog.

#### Exercise 4.2: Get the item sender ####
1. In **Home.html**, locate the "Exercise: Get sender" section (commented) and add the following HTML piece inside the **div** (section) tags. This is an Office UI Fabric styled button. 
    ```html
    <button id="get-sender" class="ms-Button ms-Button--primary">
        <span class="ms-Button-label">Get sender</span>
    </button>
    
    ```
2. In **Home.js**, add an event handler (below the initialization of the Office UI Fabric components, in the **ready** function) for the click event of the button:
    ```js
    $('#get-sender').click(getSender);
    
    ```
3. In **Home.js**, add the following function to get the item sender (below the **Office.initialize** function):
    ```js
    // Get the item sender and display it
    function getSender() {
        var _item = Office.context.mailbox.item;
        var sender;

        // Check if the item is a message or appointment
        // in order to determine which property that contains
        // the sender
        if (_item.itemType === Office.MailboxEnums.ItemType.Message) {
            sender = _item.from;
        }
        else if (_item.itemType === Office.MailboxEnums.ItemType.Appointment) {
            sender = _item.organizer;
        }

        // Show data
        showDataDialog('Sender', sender.displayName);
    }
    ```
4. Launch your mail add-in and test your work by clicking the **Get sender** button. When the button is clicked, the function will be executed; displaying the item sender using the data dialog.

#### Exercise 4.3: Get the item body ####
1. In **Home.html**, locate the "Exercise: Get body" section (commented) and add the following HTML piece inside the **div** (section) tags. This is an Office UI Fabric styled button. 
    ```html
    <button id="get-body" class="ms-Button">
        <span class="ms-Button-label">Get body</span>
    </button>
    
    ```
2. In **Home.js**, add an event handler (below the initialization of the Office UI Fabric components, in the **ready** function) for the click event of the button:
    ```js
    $('#get-body').click(getBody);
    
    ```
3. In **Home.js**, add the following function to get the item body (below the **Office.initialize** function):
    ```js
    // Get the body of the item and display it as 
    // plain text
    function getBody() {
        var _item = Office.context.mailbox.item;
        var body = _item.body;

        // Get the body asynchronous as text
        body.getAsync(Office.CoercionType.Text, function (asyncResult) {
            if (asyncResult.status !== Office.AsyncResultStatus.Succeeded) {
                // TODO: Handle error
            }
            else {
                // Show data
                showDataDialog('Body', asyncResult.value.trim());
            }
        });
    }
    ```
4. Getting the item body is an asynchronous function that requires a minimum mailbox requirement set version of 1.3. There are different ways of ensuring that your user has at least version 1.3, a good way is to set it in the manifest.       
         
   In the manifest project **Read-Mode-Outlook-Add-in**, double-click the **Read-Mode-Outlook-Add-inManifest** file. This will open the manifest editor.      
   ![](https://raw.githubusercontent.com/simonjaeger/OfficeDev-HOL/master/Read-Mode-Outlook-Add-in/Images/OutlookAddinManifest.png)
5. In the **General** tab section, find the **Mailbox requirement set** property and set it to **1.3**.       
   ![](https://raw.githubusercontent.com/simonjaeger/OfficeDev-HOL/master/Read-Mode-Outlook-Add-in/Images/RequirementSet.png)
4. Launch your mail add-in and test your work by clicking the **Get body** button. When the button is clicked, the function will be executed; displaying the item body using the data dialog.            
   ![](https://raw.githubusercontent.com/simonjaeger/OfficeDev-HOL/master/Read-Mode-Outlook-Add-in/Images/LaunchedReadMailAddin4.png)

# Wrap up  #
View the source code files included in this hands-on lab for a final reference of how your code should be structured (if needed). You should now have grasped an understanding of a few possibilities of interacting with the Office context (a mailbox item in this case). In addition, you have also seen some of the styles and components included in the Office UI Fabric.

# More Resources #
- Discover Office development: <https://msdn.microsoft.com/en-us/office/>
- Learn more about Office UI Fabric: <http://dev.office.com/fabric/>
- Developing Outlook add-ins – where to integrate and what you can do: <http://simonjaeger.com/developing-outlook-add-ins-where-to-integrate-and-what-you-can-do/>