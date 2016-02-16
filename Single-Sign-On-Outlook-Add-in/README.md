# Hands-on Lab: Single Sign-On Outlook Add-in #

With the new application model for Office comes a brand new way of extending Office with your own functionality - using the tools and dev stacks that we already know and love. 

In mail add-ins, we have access to a great functionality within the APIs to get an identity token. This identity token represents the Office 365 user and is key when it comes to building a single sign-on experience for your users. If you would like to learn more about what’s inside the identity token – head to: http://simonjaeger.com/dissecting-and-validating-the-exchange-identity-token/

Why would you want to do this? Imagine if your add-in could sign in the user as soon as the add-in starts in Outlook – without having to ask for credentials every time. Not only is this behavior annoying from a user standpoint, but it’s also possible to overcome this because we have access to the identity token.

This hands-on lab demonstrates how to implement a single sign-on experience within an Outlook add-in. In addition, Different styles and components from the Office UI Fabric library are used throughout this Office add-in. 
The objective is to get familiar with some of the possiblities that we have when building Excel add-ins.

The hands-on lab is divided into multiple exercises and should be followed in a chronological order. These are the included exercises:

* [1.1 Create the project](#exercise-11-create-the-project)
* [1.2 Edit the manifest](#exercise-12-edit-the-manifest)
* [1.3 Launch the project](#exercise-13-launch-the-project)
* [2.1 Clean up the project](#exercise-21-clean-up-the-project)
* [2.2 Add Office UI Fabric](#exercise-22-add-office-ui-fabric)
* [2.3 Add the base](#exercise-23-add-the-base-css--html)

Short of time and just want the final sample? Clone this repository (```git clone https://github.com/simonjaeger/OfficeDev-HOL.git```) and open the solution file: **Single-Sign-On-Outlook-Add-in\\Source\\Single-Sign-On-Outlook-Add-in.sln**.           
    
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
Single-Sign-On-Outlook-Add-in | Simon Jäger (**Microsoft**)

### Version history ###
Version  | Date | Comments
---------| -----| --------
1.0  | February 16th 2016 | Initial release

### Disclaimer ###
**THIS CODE IS PROVIDED *AS IS* WITHOUT WARRANTY OF ANY KIND, EITHER EXPRESS OR IMPLIED, INCLUDING ANY IMPLIED WARRANTIES OF FITNESS FOR A PARTICULAR PURPOSE, MERCHANTABILITY, OR NON-INFRINGEMENT.**

----------

# Exercises #
#### Exercise 1.1: Create the project ####
The first thing that we need to do is to create the project itself. Make sure that you have installed all of the required prerequisites before launching Visual Studio 2015. 

1. Click **File**, **New** and finally the **Project** button.
2. In **Templates**, select **Visual C#**, **Office/SharePoint** and then **Office Add-ins**. This will list the Office add-in templates, choose **Office Add-in**. 
3. Name your project **"Single-Sign-On-Outlook-Add-in"** and click the **OK** button to continue. 
   ![](https://raw.githubusercontent.com/simonjaeger/OfficeDev-HOL/master/Single-Sign-On-Outlook-Add-in/Images/NewProject.png)
4. Next up Visual Studio 2015 will need a bit more information about what you are going to create - in order to set up the required files. Your next step is to decide which type of Office add-in that you want to create. Depending on what you pick, your Office add-in will run in different Office applications and contexts. 
   
   For this hands-on lab, we will create a mail add-in - this means that our Office add-in will run in in Outlook as a view beside the Office context (e.g. a message or appointment). Select **Mail** and click on **Next**. 
   ![](https://raw.githubusercontent.com/simonjaeger/OfficeDev-HOL/master/Read-Mode-Outlook-Add-in/Images/MailAddinType.png)
5. Finally we need to choose the supported modes for our mail add-in. This means that we are defining the contexts that our mail add-in can run within; read, compose or both. If you choose **Read form**, the mail add-in will be able to run when a user is viewing a mailbox item. In **Compose form**, the mail add-in can run when a user is creating or editing a mailbox item. 
   
   In our case, select **Read form** for **Email message**. Deselect everything else to create a read mode mail add-in. Click **Finish** to complete the wizard.
   ![](https://raw.githubusercontent.com/simonjaeger/OfficeDev-HOL/master/Single-Sign-On-Outlook-Add-in/Images/ReadMailAddin.png)
6. Using the information you specified in the wizard, Visual Studio 2015 will configure your project. Have a look in the **Solution Explorer** and find your two new projects in the **Single-Sign-On-Outlook-Add-in** solution. 
   
   **Single-Sign-On-Outlook-Add-in:** This is your manifest project, containing the XML manifest file. This is basically a representation of the information you just specified while creating your Office add-in project. 
   
   **Single-Sign-On-Outlook-Add-inWeb:** This is your web project for the Office add-in. This contains all of the different source files that makes up your Outlook add-in. We will make quite a few adjustments to this structure as we continue.
   ![](https://raw.githubusercontent.com/simonjaeger/OfficeDev-HOL/master/Single-Sign-On-Outlook-Add-in/Images/SolutionExplorer.png)
   
   You've now created the basic structure for a mail add-in running in Outlook. 

#### Exercise 1.2: Edit the manifest ####
We need to make sure that we understand the manifest file. This file is essential for your add-in; it tells Office where everything is hosted (locally throughout this hands-on lab) and where it can be launched. So let's open it and edit the manifest file.

1. In the manifest project **Single-Sign-On-Outlook-Add-in**, double-click the **Single-Sign-On-Outlook-Add-inManifest** file. This will open the manifest editor.      
   ![](https://raw.githubusercontent.com/simonjaeger/OfficeDev-HOL/master/Single-Sign-On-Outlook-Add-in/Images/OutlookAddinManifest.png)
2. In the **General** tab section, find and edit the **Display name** and **Provider name** to anything you'd like.      
   ![](https://raw.githubusercontent.com/simonjaeger/OfficeDev-HOL/master/Single-Sign-On-Outlook-Add-in/Images/GeneralAddinManifest.png)
3. In the **Read Form** tab section, find the **Activation** part. This is what determines the rules for potential activation of your mail add-in. By default, **Item is a message** should be included. 
4. Scroll down and pay attention to the **Source location** property. This points to a specific file in your web project (**Single-Sign-On-Outlook-Add-inWeb**). When launching your mail add-in, this page will be the first thing that gets loaded and displayed. 
5. Below the **Source location** property, you will find the **Requested height (px)** property. This a way for your mail add-in to ask for a certain height in pixels when displayed within Outlook. Be aware though, it doesn't mean that it will be granted. Change this to **300** - as we want a bit more height for our mail add-in to display the sign-in form. 
   ![](https://raw.githubusercontent.com/simonjaeger/OfficeDev-HOL/master/Single-Sign-On-Outlook-Add-in/Images/ReadFormAddinManifest.png)

#### Exercise 1.3: Launch the project ####
Before we launch our mail add-in we should validate that our start actions are proper.

1. Select the manifest project; **Single-Sign-On-Outlook-Add-in** in the **Solution Explorer**.                                     
   ![](https://raw.githubusercontent.com/simonjaeger/OfficeDev-HOL/master/Single-Sign-On-Outlook-Add-in/Images/SelectManifestProject.png)
2. In the **Properties** window, set **Start Action** to Office Desktop Client. 
4. Set **Web Project** to your web project; **Single-Sign-On-Outlook-Add-inWeb**.
   ![](https://raw.githubusercontent.com/simonjaeger/OfficeDev-HOL/master/Single-Sign-On-Outlook-Add-in/Images/StartActions.png)
5. To launch the project, open on the **Debug** menu at the top of Visual Studio 2015 and click on the **Start Debugging** button. You can also click the **Start** button in your toolbar or use the **{F5}** keyboard shortcut.            
   ![](https://raw.githubusercontent.com/simonjaeger/OfficeDev-HOL/master/Excel-Add-in/Images/StartProject.png)
6. When launching your mail add-in for the first time, Visual Studoo 2015 needs to install the manifest file. This is where you should use your Office 365 Developer Tenant (if you haven't signed up for one yet, get yours for free at <http://dev.office.com/devprogram>). Enter the credentials of a user (**[username]@[your domain].onmicrosoft.com**) belonging to your Office 365 Developer Tenant and click on the **Connect** button.            
   ![](https://raw.githubusercontent.com/simonjaeger/OfficeDev-HOL/master/Read-Mode-Outlook-Add-in/Images/ConnectToExchange.png)
7. Once Outlook has launched, you'll notice that your mail add-in doesn't start right away. We need to start it manually. Select a message in your mailbox (send yourself one if needed) and click on the **Single-Sign-On-Outlook-Add-in** above it. Once your mail add-in has launched, you can explore the functionality that comes right out of the box with the Visual Studio 2015 template.            
   ![](https://raw.githubusercontent.com/simonjaeger/OfficeDev-HOL/master/Single-Sign-On-Outlook-Add-in/Images/LaunchedSSOMailAddin.png)
8. Finally, stop debugging by opening the **Debug** menu at the top of Visual Studio 2015 and click on the **Stop Debugging** button. You can also click the **Stop** button in your toolbar.            
   ![](https://raw.githubusercontent.com/simonjaeger/OfficeDev-HOL/master/Word-Add-in/Images/StopDebugging.png)

#### Exercise 2.1: Clean up the project ####
While the default styling that comes along with the Visual Studio 2015 template for Office add-ins does its job - leveraging the features of the Office UI Fabric can be fantastic. It's a UI toolkit made specifically for building Office and Office 365 experiences, so it will certainly help us out here.

The Office UI Fabric library comes with everything from styling, components to animations. The majority of the library can be references via a CDN. The heavier parts needs to be downloaded and added to the project itself. We will go through both of these approaches. 

Our first task is to clean up the project, and remove the default styling and setup.

1. Remove the **Content** and **Images** folders from the web project. You can do this by right-clicking these folders in the **Solution Explorer** and choosing the **Delete** option.                                    
    ![](https://raw.githubusercontent.com/simonjaeger/OfficeDev-HOL/master/Word-Add-in/Images/DeleteFolders.png)
2. In your **Solution Explorer**, find the **Home.html** file - which is the startup page for your mail add-in. Remove everything inside the **body** tags. 
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
4. In **App.js**, remove the **initialize()** function defined on the **app** object, as this will not be used:            
    ```js
    var app = (function () {
        "use strict";
    
        var app = {};
        return app;
    })();
    ```
5. In **Home.js**, remove the **displayItemDetails()** function and the call to **app.initialize()**. We are remaking the structure of the mail add-in, these will no longer be used. You should end up with this:            
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
6. In **App.css**, remove everything, leaving you with an empty file.
7. In **Home.css**, remove everything, leaving you with an empty file.

#### Exercise 2.2: Add Office UI Fabric ####
1. In **Home.html**, add two CSS references to the CDN for Office UI Fabric inside the **head** tags. Add them before the CSS reference to **"../App.css"**.
    ```html
    <link rel="stylesheet" href="https://appsforoffice.microsoft.com/fabric/1.0/fabric.min.css">
    <link rel="stylesheet" href="https://appsforoffice.microsoft.com/fabric/1.0/fabric.components.min.css">
    
    ```
    
2. Some components in the Office UI Fabric library require some additional JavaScript to function. In our case, we will use a Spinner component that needs this. **Download** the JavaScript file for this component (**Spinner.js**) at <hhttps://raw.githubusercontent.com/OfficeDev/Office-UI-Fabric/master/src/components/Spinner/Spinner.js> or get it by browsing the files included in this hands-on lab. 
3. Add the **Spinner.js** file to your **Scripts** folder in the **Solution Explorer**. You can do this by right-clicking the **Scripts** folder and choosing **Add Existing Item...**.                                    
    ![](https://raw.githubusercontent.com/simonjaeger/OfficeDev-HOL/master/Word-Add-in/Images/AddExisting.png)
4. In **Home.html**, reference the **Spinner.js** file by adding the following line inside the **head** tags. Be sure to add it after the reference to **"../../Scripts/jquery-1.9.1.js"**.                      
    ```html
    <script src="../../Scripts/Spinner.js" type="text/javascript"></script>
    
    ```
5. In **Home.js**, add the following line inside the **document.ready** function of your page.             
    ```js
    if (typeof fabric !== "") {
        if ('Spinner' in fabric) {
            var element = document.querySelector('.ms-Spinner');
            if (element) {
                var component = new fabric['Spinner'](element);
            }
        }
    }
    ```  
    This will use the functionality within the **Spinner.js** file and initialize any Spinner components that we add to the view. 
    
6. Your **Home.html** file should now look like this: 
    ```html
    <!DOCTYPE html>
    <html>
    <head>
        <meta charset="UTF-8" />
        <meta http-equiv="X-UA-Compatible" content="IE=Edge" />
        <title></title>
        <script src="../../Scripts/jquery-1.9.1.js" type="text/javascript"></script>

        <script src="https://appsforoffice.microsoft.com/lib/1/hosted/office.js" type="text/javascript"></script>

        <script src="../../Scripts/Spinner.js" type="text/javascript"></script>

        <link rel="stylesheet" href="https://appsforoffice.microsoft.com/fabric/1.0/fabric.min.css">
        <link rel="stylesheet" href="https://appsforoffice.microsoft.com/fabric/1.0/fabric.components.min.css">

        <link href="../App.css" rel="stylesheet" type="text/css" />
        <script src="../App.js" type="text/javascript"></script>

        <link href="Home.css" rel="stylesheet" type="text/css" />
        <script src="Home.js" type="text/javascript"></script>
    </head>
    <body>

    </body>
    </html>

    ```  
    Your **Home.js** file should look like this:
    ```js
    (function () {
        "use strict";

        // The Office initialize function must be run each time a new page is loaded
        Office.initialize = function (reason) {
            $(document).ready(function () {
                // Initialize Office UI Fabric components (spinner)
                if (typeof fabric !== "") {
                    if ('Spinner' in fabric) {
                        var element = document.querySelector('.ms-Spinner');
                        if (element) {
                            var component = new fabric['Spinner'](element);
                        }
                    }
                }

            });
        };
    })();
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
        margin-bottom: 15px;
    }
    ``` 
2. In **Home.html**, add the following chunk of HTML inside the **body** tags. This will set the stage for the next array of exercises to come. 
    ```html
    <!-- Header -->
    <div id="header" class="ms-bgColor-themePrimary">
        <h2 class="ms-font-xl ms-fontWeight-semibold ms-fontColor-white">
            HOL: Single Sign-On Outlook Add-in
        </h2>
    </div>
    <div id="content">
        <div class="ms-Grid">
            <div class="ms-Grid-row">
                <div class="ms-Grid-col ms-u-sm5 ms-u-md5" style="padding-left: 0px;">
                    <p class="ms-font-l ms-fontWeight-semibold section-title">Exercise: SSO</p>
                    <!-- Introduction -->
                    <p class="ms-font-m ms-fontColor-neutralSecondary">
                        This sample demonstrates how to implement a single sign-on experience within an Outlook add-in.
                        <br />
                        <br />
                        Enter the credentials of an existing user in your user service before signing in.
                    </p>
                </div>

                <!-- Exercise Section: SSO -->
                <div class="ms-Grid-col ms-u-sm7 ms-u-md7" style="padding-top: 15px;">
                    <!-- Exercise: Username -->

                    <!-- Exercise: Password -->

                    <div class="ms-Grid">
                        <div class="ms-Grid-row">
                            <div class="ms-Grid-col ms-u-sm7 ms-u-md7">
                                <!-- Exercise: ChoiceField -->

                            </div>
                            <div class="ms-Grid-col ms-u-sm5 ms-u-md5">
                                <!-- Exercise: Sign in button -->

                                <!-- Exercise: Spinner -->

                            </div>
                        </div>
                    </div>
                </div>
            </div>
        </div>
    </div>

    <!-- Exercise: Data dialog -->

    ``` 
3. Launch your mail add-in to display the new styling. We will add more interactive components in the different sections (inside the recently added HTML piece).            
   ![](https://raw.githubusercontent.com/simonjaeger/OfficeDev-HOL/master/Single-Sign-On-Outlook-Add-in/Images/LaunchedSSOMailAddin2.png)


#### Exercise 4.1: Add the input fields ####
1. In **Home.html**, locate the "Exercise: Username" comment and add the following HTML piece below it. This is an Office UI Fabric styled input field. 
    ```html
    <div class="ms-TextField">
        <label class="ms-Label">Username</label>
        <input id="username" class="ms-TextField-field" type="text" disabled>
        <span class="ms-TextField-description">This should be an entry in your user service</span>
    </div>
    
    ```
2. In **Home.html**, locate the "Exercise: Password" comment and add the following HTML piece below it. This is an Office UI Fabric styled input field. 
    ```html
    <div class="ms-TextField">
        <label class="ms-Label">Password</label>
        <input id="password" class="ms-TextField-field" type="password" disabled>
    </div>
    
    ```
2. In **Home.html**, locate the "Exercise: ChoiceField" comment and add the following HTML piece below it. This is an Office UI Fabric styled input field (checkbox). 
    ```html
    <div class="ms-ChoiceField">
        <input id="demo-checkbox-unselected" class="ms-ChoiceField-input" type="checkbox" checked disabled>
        <label for="demo-checkbox-unselected" class="ms-ChoiceField-field">
            <span class="ms-Label">Keep me signed in</span>
        </label>
    </div>
    
    ```
2. In **Home.js**, add some default values to our input fields (below the initialization of the Office UI Fabric components, in the **ready** function) when the mail add-in is initialized:
    ```js
    $('#username').val('#Office365Dev');
    $('#password').val('#Office365Dev');
    
    ```
4. Launch your mail add-in and view your work. You should see the new input fields that we will use as the sign-in form. You will see that the input fields are disabled, we will enable them later on as part of the initalization logic.            
   ![](https://raw.githubusercontent.com/simonjaeger/OfficeDev-HOL/master/Single-Sign-On-Outlook-Add-in/Images/LaunchedSSOMailAddin3.png)

#### Exercise 4.2: Add the Spinner ####
1. In **Home.html**, locate the "Exercise: Spinner" comment and add the following HTML piece below it. This is an Office UI Fabric styled input field. 
    ```html
    <div id="spinner" class="ms-Spinner ms-Spinner--large">
    </div>
    
    ```
2. In **Home.css**, add the following CSS piece. This will place the Spinner within the form and customize the color.
    ```css
    #spinner {
        float: right;
        width: 32px;
        height: 32px;
        margin-right: -3px;
    }

    .ms-Spinner-circle {
        background-color: #0078D7;
    }
    
    ```
4. Launch your mail add-in and view your work. You should be able to see the Office UI Spinner. We will use this when we are waiting for a response from the backend.             
   ![](https://raw.githubusercontent.com/simonjaeger/OfficeDev-HOL/master/Single-Sign-On-Outlook-Add-in/Images/LaunchedSSOMailAddin4.png)

#### Exercise 4.3: Add the dialog ####

1. In **Home.html**, locate the "Exercise: Data dialog" comment and add the following HTML piece below it. This is an Office UI Fabric styled dialog. 
    ```html
    <div id="dialog" class="ms-Dialog ms-Dialog--lgHeader">
        <div class="ms-Overlay"></div>
        <div class="ms-Dialog-main">
            <button class="ms-Dialog-button ms-Dialog-button--close">
                <i class="ms-Icon ms-Icon--x"></i>
            </button>
            <div class="ms-Dialog-header">
                <p id="dialog-title" class="ms-Dialog-title">
                    <!-- Dialog title -->
                </p>
            </div>
            <div class="ms-Dialog-inner">
                <div class="ms-Dialog-content">
                    <p id="dialog-text" class="ms-Dialog-subText" style="max-height: 95px; overflow-y: hidden;">
                        <!-- Dialog text -->
                    </p>
                </div>
                <div class="ms-Dialog-actions">
                    <div class="ms-Dialog-actionsRight">
                        <button id="dialog-ok" class="ms-Dialog-action ms-Button ms-Button--primary">
                            <span class="ms-Button-label">OK</span>
                        </button>
                    </div>
                </div>
            </div>
        </div>
    </div>
    
    ```
2. Launch your mail add-in and test your work. You should find a dialog covering up most of your display area.             
   ![](https://raw.githubusercontent.com/simonjaeger/OfficeDev-HOL/master/Single-Sign-On-Outlook-Add-in/Images/LaunchedSSOMailAddin5.png)
3. In **Home.js**, add the following event handler (below the initialization of input field values) for the dialog. This will allow us to close it.
    ```js
    $('#dialog-ok').click(hideDialog);
    
    ```
4. In **Home.js**, add the following functions to show and close the dialog:
    ```js
    // Show the dialog
    function showDialog(title, text) {
        // Set the dialog title and text
        $('#dialog-title').text(title);
        $('#dialog-text').text(text);
        $('#dialog').show();
    }

    // Hide the dialog
    function hideDialog() {
        $('#dialog').hide();
    }
    ```
5. Launch your mail add-in and test your work. You should be able to close the dialog using the **OK** button.
6. In **Home.css**, add the following CSS piece to hide the dialog when your mail add-in has launched.
    ```css
    #dialog {
        display: none;
    }
    
    ```

#### Exercise 5.1: Add the authentication logic ####
In order to authenticate and create a single sign-on experience, we need to perform a couple of things:
- **Add-in:** Get the identity token of the current Office 365 user.
- **Add-in:** Send the identity token to the Web API and check if the identity token has already been mapped.
- **Web API:** If a mapping is found (no user credentials needed)
    -  **Add-in:** Return the user data and sign in automatically (single sign-on).  
- **Web API:** If a mappiung is not found (user credentials needed)
    - **Add-in:** Let the user sign in using their credentials.
    - **Add-in:** Send the identity token and user credentials to the Web API.
    - **Web API:** If the provided credentials are correct, map the identity token with the user in the Web API.
    - **Web API:** Notify the mail add-in.
    - **Add-in:** Reload the mail add-in and experience single sign-on.

We need to implement two parts to achieve the above; the front-end (add-in) and backend (Web API). Let's begin with the add-in side of things.

1. In **Home.js**, add the following function to get the current ID token (for the Office 365 user) and send it along with optional credentials to the Web API. 
    ```js
    // Try to authenticate, if credentials for a user is provided - we
    // will try to map it with the Office 365 user in the backend
    function authenticate(credentials, callback) {
        Office.context.mailbox.getUserIdentityTokenAsync(function (asyncResult) {
            if (asyncResult.status !== Office.AsyncResultStatus.Succeeded) {
                // TODO: Handle error
                callback();
            }
            else {
                var token = asyncResult.value;

                // POST the credentials to the Web API
                $.ajax({
                    type: 'POST',
                    url: '../../api/sso',
                    data: JSON.stringify({
                        identityToken: token,
                        hostUri: window.location.href.split('?')[0],
                        credentials: credentials
                    }),
                    contentType: 'application/json;charset=utf-8'
                }).done(function (response) {
                    // TODO: Validate response
                    callback(response);
                }).fail(function (error) {
                    // TODO: Handle error
                    callback();
                });
            }
        });
    }
    
    ```

#### Exercise 5.2: Add the sign in button ####
1. In **Home.html**, locate the "Exercise: Sign in button" comment and add the following HTML piece below it. This is an Office UI Fabric styled input field. 
    ```html
    <button id="sign-in" class="ms-Button ms-Button--primary" disabled>
        <span class="ms-Button-label">Sign in</span>
    </button>
    
    ```
2. In **Home.css**, add the following CSS piece. This will place the sign in button within the form and hide it.
    ```css
    #sign-in {
        display: none;
        float: right;
    }
    
    ```
2. In **Home.js**, add an event handler for the click event of the sign in button:
    ```js
    $('#sign-in').click(signIn);
    
    ```
3. In **Home.js**, add the following function to perform the sign in logic (described in Exercise 5.1).
    ```js
    // Sign in using the provided credentials
    function signIn() {
        // Show spinner and hide button
        $('#sign-in').hide();
        $('#spinner').show();

        // Disable input fields
        $('#username').attr('disabled', 'disabled');
        $('#password').attr('disabled', 'disabled');

        // Get credentials
        var credentials = {
            username: $('#username').val(),
            password: $('#password').val(),
        }

        // Authenticate
        authenticate(credentials, function (response) {
            if (!response) {
                // Hide spinner and show button
                $('#sign-in').show();
                $('#spinner').hide();

                // Enable input fields
                $('#username').removeAttr('disabled');
                $('#password').removeAttr('disabled');

                // Display error
                showDialog('Oops!', 'Something happened... make sure that ' +
                    'the credentials are valid (for an entry in the user service).');
            }
            else {
                // Reload page and try SSO
                location.reload();
            }
        });
    }
    ```
    
#### Exercise 5.3: Enable the sign in button ####
1. In **Home.js**, add the following code in the **document.ready** function (below the event handlers) to perform the single sign-on logic when the mail add-in is initialized. This will enable the sign in form if single sign-on is not possible (no mapping of the user in the Web API).
    ```js
    // Authenticate silently (without credentials)
    authenticate(null, function (response) {
        // Hide spinner and show button
        $('#sign-in').show();
        $('#spinner').hide();

        if (!response) {
            // Enable sign-in
            $('#sign-in').removeAttr('disabled');
            $('#username').removeAttr('disabled');
            $('#password').removeAttr('disabled');
        }
        else {
            // Display user data
            showDialog('Hi developer!', 'Office 365 user ' + '(' +
                Office.context.mailbox.userProfile.emailAddress + ') ' +
                'has automatically signed in as user: ' +
                response.displayName + ' (' +
                response.credentials.username + ').');
        }
    });
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

        <script src="../../Scripts/Spinner.js"></script>

        <link href="../App.css" rel="stylesheet" type="text/css" />
        <script src="../App.js" type="text/javascript"></script>

        <link href="Home.css" rel="stylesheet" type="text/css" />
        <script src="Home.js" type="text/javascript"></script>
    </head>
    <body>
        <!-- Header -->
        <div id="header" class="ms-bgColor-themePrimary">
            <h2 class="ms-font-xl ms-fontWeight-semibold ms-fontColor-white">
                HOL: Single Sign-On Outlook Add-in
            </h2>
        </div>
        <div id="content">
            <div class="ms-Grid">
                <div class="ms-Grid-row">
                    <div class="ms-Grid-col ms-u-sm5 ms-u-md5" style="padding-left: 0px;">
                        <p class="ms-font-l ms-fontWeight-semibold section-title">Exercise: SSO</p>
                        <!-- Introduction -->
                        <p class="ms-font-m ms-fontColor-neutralSecondary">
                            This sample demonstrates how to implement a single sign-on experience within an Outlook add-in.
                            <br />
                            <br />
                            Enter the credentials of an existing user in your user service before signing in.
                        </p>
                    </div>

                    <!-- Exercise Section: SSO -->
                    <div class="ms-Grid-col ms-u-sm7 ms-u-md7" style="padding-top: 15px;">
                        <!-- Exercise: Username -->
                        <div class="ms-TextField">
                            <label class="ms-Label">Username</label>
                            <input id="username" class="ms-TextField-field" type="text" disabled>
                            <span class="ms-TextField-description">This should be an entry in your user service</span>
                        </div>

                        <!-- Exercise: Password -->
                        <div class="ms-TextField">
                            <label class="ms-Label">Password</label>
                            <input id="password" class="ms-TextField-field" type="password" disabled>
                        </div>

                        <div class="ms-Grid">
                            <div class="ms-Grid-row">
                                <div class="ms-Grid-col ms-u-sm7 ms-u-md7">
                                    <!-- Exercise: ChoiceField -->
                                    <div class="ms-ChoiceField">
                                        <input id="demo-checkbox-unselected" class="ms-ChoiceField-input" type="checkbox" checked disabled>
                                        <label for="demo-checkbox-unselected" class="ms-ChoiceField-field">
                                            <span class="ms-Label">Keep me signed in</span>
                                        </label>
                                    </div>
                                </div>
                                <div class="ms-Grid-col ms-u-sm5 ms-u-md5">
                                    <!-- Exercise: Sign in button -->
                                    <button id="sign-in" class="ms-Button ms-Button--primary" disabled>
                                        <span class="ms-Button-label">Sign in</span>
                                    </button>

                                    <!-- Exercise: Spinner -->
                                    <div id="spinner" class="ms-Spinner ms-Spinner--large">
                                    </div>
                                </div>
                            </div>
                        </div>
                    </div>
                </div>
            </div>
        </div>

        <!-- Exercise: Data dialog -->
        <div id="dialog" class="ms-Dialog ms-Dialog--lgHeader">
            <div class="ms-Overlay"></div>
            <div class="ms-Dialog-main">
                <button class="ms-Dialog-button ms-Dialog-button--close">
                    <i class="ms-Icon ms-Icon--x"></i>
                </button>
                <div class="ms-Dialog-header">
                    <p id="dialog-title" class="ms-Dialog-title">
                        <!-- Dialog title -->
                    </p>
                </div>
                <div class="ms-Dialog-inner">
                    <div class="ms-Dialog-content">
                        <p id="dialog-text" class="ms-Dialog-subText" style="max-height: 95px; overflow-y: hidden;">
                            <!-- Dialog text -->
                        </p>
                    </div>
                    <div class="ms-Dialog-actions">
                        <div class="ms-Dialog-actionsRight">
                            <button id="dialog-ok" class="ms-Dialog-action ms-Button ms-Button--primary">
                                <span class="ms-Button-label">OK</span>
                            </button>
                        </div>
                    </div>
                </div>
            </div>
        </div>
    </body>
    </html>


    ```  
    Your **Home.js** file should now look like this:
    ```js
    (function () {
        "use strict";

        // The Office initialize function must be run each time a new page is loaded
        Office.initialize = function (reason) {
            $(document).ready(function () {
                // Initialize Office UI Fabric components (spinner)
                if (typeof fabric !== "") {
                    if ('Spinner' in fabric) {
                        var element = document.querySelector('.ms-Spinner');
                        if (element) {
                            var component = new fabric['Spinner'](element);
                        }
                    }
                }

                // Initialize input fields
                $('#username').val('#Office365Dev');
                $('#password').val('#Office365Dev');

                // Add event handlers
                $('#sign-in').click(signIn);
                $('#dialog-ok').click(hideDialog);

                // Authenticate silently (without credentials)
                authenticate(null, function (response) {
                    // Hide spinner and show button
                    $('#sign-in').show();
                    $('#spinner').hide();

                    if (!response) {
                        // Enable sign-in
                        $('#sign-in').removeAttr('disabled');
                        $('#username').removeAttr('disabled');
                        $('#password').removeAttr('disabled');
                    }
                    else {
                        // Display user data
                        showDialog('Hi developer!', 'Office 365 user ' + '(' +
                            Office.context.mailbox.userProfile.emailAddress + ') ' +
                            'has automatically signed in as user: ' +
                            response.displayName + ' (' +
                            response.credentials.username + ').');
                    }
                });
            });
        };

        // Try to authenticate, if credentials for a user is provided - we
        // will try to map it with the Office 365 user in the backend
        function authenticate(credentials, callback) {
            Office.context.mailbox.getUserIdentityTokenAsync(function (asyncResult) {
                if (asyncResult.status !== Office.AsyncResultStatus.Succeeded) {
                    // TODO: Handle error
                    callback();
                }
                else {
                    var token = asyncResult.value;

                    // POST the credentials to the Web API
                    $.ajax({
                        type: 'POST',
                        url: '../../api/sso',
                        data: JSON.stringify({
                            identityToken: token,
                            hostUri: window.location.href.split('?')[0],
                            credentials: credentials
                        }),
                        contentType: 'application/json;charset=utf-8'
                    }).done(function (response) {
                        // TODO: Validate response
                        callback(response);
                    }).fail(function (error) {
                        // TODO: Handle error
                        callback();
                    });
                }
            });
        }

        // Sign in using the provided credentials
        function signIn() {
            // Show spinner and hide button
            $('#sign-in').hide();
            $('#spinner').show();

            // Disable input fields
            $('#username').attr('disabled', 'disabled');
            $('#password').attr('disabled', 'disabled');

            // Get credentials
            var credentials = {
                username: $('#username').val(),
                password: $('#password').val(),
            }

            // Authenticate
            authenticate(credentials, function (response) {
                if (!response) {
                    // Hide spinner and show button
                    $('#sign-in').show();
                    $('#spinner').hide();

                    // Enable input fields
                    $('#username').removeAttr('disabled');
                    $('#password').removeAttr('disabled');

                    // Display error
                    showDialog('Oops!', 'Something happened... make sure that ' +
                        'the credentials are valid (for an entry in the user service).');
                }
                else {
                    // Reload page and try SSO
                    location.reload();
                }
            });
        }

        // Show the dialog
        function showDialog(title, text) {
            // Set the dialog title and text
            $('#dialog-title').text(title);
            $('#dialog-text').text(text);
            $('#dialog').show();
        }

        // Hide the dialog
        function hideDialog() {
            $('#dialog').hide();
        }
    })();
    ``` 

3. Launch your mail add-in and test your work. Quickly after launching, your sign in form should be enabled. If you try to sign in, you will be prompted by the dialog that we created - saying that an error occurred. Which is true as we haven't created our Web API yet, we will do that next.
   ![](https://raw.githubusercontent.com/simonjaeger/OfficeDev-HOL/master/Single-Sign-On-Outlook-Add-in/Images/LaunchedSSOMailAddin6.png)

#### Exercise 6.1: Add the Web API Controller ####
Now let's create the backend for our mail add-in. We will create a Web API to respond and validate any requests for single sign-on. To get started, we should add a Web API Controller to the project - this will trigger Visual Studio to download the required NuGet packages.

1. Select the web project; **Read-Mode-Outlook-Add-inWeb** in the **Solution Explorer**.       
   ![](https://raw.githubusercontent.com/simonjaeger/OfficeDev-HOL/master/Single-Sign-On-Outlook-Add-in/Images/SelectWebProject.png)
2. Right-click and choose **Add New Folder**, name it **"Controllers"**. 
3. Right-click on the **Controllers** folder and choose **Web API Controller Class (v2.1)**. Name it **"SSOController"** and click on the **OK** button.        
   ![](https://raw.githubusercontent.com/simonjaeger/OfficeDev-HOL/master/Single-Sign-On-Outlook-Add-in/Images/SSO.png) 
4. In **SSOController.cs**, remove every method leaving you with an empty class:  
    ```csharp
    using System.Web.Http;

    namespace Single_Sign_On_Outlook_Add_inWeb.Controllers
    {
        public class SSOController : ApiController
        {
        }
    }
    ```

#### Exercise 6.2: Add the Web API configuration ####
We need to configure a router for the Web API, this means that we define how you should structure the URLs to access the Controllers and its operations. It's very straight forward. 

1. Select the web project; **Read-Mode-Outlook-Add-inWeb** in the **Solution Explorer**.
2. Right-click and choose **Add New Folder**, name it **"App_Start"**. 
3. Right-click on the **App_Start** folder and choose **Add Class...**, name it **"WebApiConfig"**.        
   ![](https://raw.githubusercontent.com/simonjaeger/OfficeDev-HOL/master/Single-Sign-On-Outlook-Add-in/Images/WebApiConfig.png)
4. In **WebApiConfig.cs**, remove everything and add the following code piece. This class will configure the route for the Web API (this determines how we call it). 
    ```csharp
    using System.Web.Http;

    namespace Single_Sign_On_Outlook_Add_inWeb.App_Start
    {
        class WebApiConfig
        {
            public static void Register(HttpConfiguration configuration)
            {
                configuration.Routes.MapHttpRoute("API Default", "api/{controller}/{id}",
                    new { id = RouteParameter.Optional });
            }
        }
    }
    
    ```
    
#### Exercise 6.3: Add the Global Application Class ####
In order to use the Web API configuration class (**WebApiConfig.cs**), we need to trigger the **Register** method upon the initialization of the web application. We can do that using a **Global Application Class** with the **Application_Start** method.

1. Select the web project; **Read-Mode-Outlook-Add-inWeb** in the **Solution Explorer**.
2. Right-click and choose **Add Global Application Class**. Name it **"Global"** and click on the **OK** button.       
   ![](https://raw.githubusercontent.com/simonjaeger/OfficeDev-HOL/master/Single-Sign-On-Outlook-Add-in/Images/Global.png)
3. In **Global.asax.cs**, find the **Application_Start** method and add the following code piece inside of it. 
    ```csharp
    // Configure Web API
    WebApiConfig.Register(GlobalConfiguration.Configuration);

    // Configure camel case for JSON responses
    var formatters = GlobalConfiguration.Configuration.Formatters;
    var jsonFormatter = formatters.JsonFormatter;
    var settings = jsonFormatter.SerializerSettings;
    settings.Formatting = Formatting.Indented;
    settings.ContractResolver = new CamelCasePropertyNamesContractResolver();
    
    ```
4. In **Global.asax.cs**, add the following using statements at the top of the file: 
    ```csharp
    using Single_Sign_On_Outlook_Add_inWeb.App_Start;
    using System.Web.Http;
    using Newtonsoft.Json;
    using Newtonsoft.Json.Serialization;
    
    ```

#### Exercise 6.4: Add the models ####
We will use three different models (objects) when building the authentication/mapping logic. These are:
- **CredentialsModel:** Represents the basic values needed to validate someone as a user (UserModel). This model is merely a demonstration of an authentication mechanism. Make sure that your usage of credentials meets the security requirements of your application. You could either implement a token flow (i.e. using one-time tokens or the OAuth protocol) for your web service or POST directly over HTTPS (with an HTML FORM) to your API.  
      
    Ideally your web service should generate a set of (short-lived) credentials only intended for the user mapping. Validate and expire those credentials in your Web API.
- **UserModel:** Represents the user in the service. It contains both a **DisplayName** property and **Credentials** property.
- **UserRequestModel:** Represents the properties (**IdentityToken**, **HostUri** and optional **Credentials**) needed to perform a mapping of an Office 365 user and a user in your service. It also containts the properties needed to do a lookup of an already mapped user.

1. Select the web project; **Read-Mode-Outlook-Add-inWeb** in the **Solution Explorer**.
2. Right-click and choose **Add New Folder**, name it **"Models"**. 
3. Right-click on the **Models** folder and choose **Add Class...**. Name it **"CredentialsModel"** and click on the **OK** button.
4. In **CredentialsModel.cs**, replace everything with the following code piece.  
    ```csharp
    namespace Single_Sign_On_Outlook_Add_inWeb.Models
    {
        public class CredentialsModel
        {
            public string Username { get; set; }
            public string Password { get; set; }
        }
    }
    
    ```
5. Right-click on the **Models** folder and choose **Add Class...**. Name it **"UserModel"** and click on the **OK** button.
6. In **UserModel.cs**, replace everything with the following code piece.  
    ```csharp
    namespace Single_Sign_On_Outlook_Add_inWeb.Models
    {
        public class UserModel
        {
            // The display name of the user
            public string DisplayName { get; set; }

            // The credentials used to verify someone as the  user
            public CredentialsModel Credentials { get; set; }
        }
    }
    
    ```
7. Right-click on the **Models** folder and choose **Add Class...**. Name it **"UserRequestModel"** and click on the **OK** button.
8. In **UserRequestModel.cs**, replace everything with the following code piece.  
    ```csharp
    namespace Single_Sign_On_Outlook_Add_inWeb.Models
    {
        public class UserRequestModel
        {
            // The raw identity token given by the mail add-in
            public string IdentityToken { get; set; }

            // The source location of the add-in
            public string HostUri { get; set; }

            // Credentials for a user in the web service - may be null.
            // Used when mapping web service users and Office 365 users.
            public CredentialsModel Credentials { get; set; }
        }
    }

    ```


#### Exercise 6.5: Add the user service interface ####
Let's create an interface to define the three different operations we need to create a single sign-on experience:
- **GetUserAsync:** Get an already mapped user in the service using the Office 365 unique user identifier (UUID).
- **GetUUIDSaltAsync:** Create or get an existing unique salt for the Office 365 unique user identifier (UUID) - used creating the UUID.
- **MapUserAsync:** Map the Office 365 unique user identifier (UUID) with a user in your service (using the provided credentials). 

1. Select the web project; **Read-Mode-Outlook-Add-inWeb** in the **Solution Explorer**.
2. Right-click and choose **Add New Folder**, name it **"Services"**. 
3. Right-click on the **Services** folder and choose **Add Class...**. Name it **"IUserService"** and click on the **OK** button.
4. In **IUserService.cs**, replace everything with the following code piece.  
    ```csharp
    using Single_Sign_On_Outlook_Add_inWeb.Models;
    using System.Threading.Tasks;

    namespace Single_Sign_On_Outlook_Add_inWeb.Services
    {
        interface IUserService
        {
            // Try to get a user service user mapped to an Office 365 user 
            // with the UUID - created by the Exchange Web Services Managed API. 
            Task<UserModel> GetUserAsync(string uuid);

            // Get a unique salt for the UUID - which is used when hashing 
            // it for user mapping. If it doesn't exist in the store - it
            // will be randomly generated. 
            Task<byte[]> GetUUIDSaltAsync(string uuid);

            // Map a user service user to an Office 365 user with the 
            // UUID. If the mapping is successful, the user is returned.
            Task<UserModel> MapUserAsync(string uuid, CredentialsModel credentials);
        }
    }

    
    ```
    

#### Exercise 6.6: Add the user service implementation ####
In order to get going and test this out - we need to implement the IUserService interface. In this exercise we will implement the three needed methods (**GetUserAsync**, **GetUUIDSaltAsync** and **MapUserAsync**) by creating a fake user service.

We will simulate a persisted solution for users, salts and mappings with simple Dictionary objects. While it will be simple to implement, it is not persisted to disk and will only be available as long as your web application is running (so keep the debugging session going to keep the data). 

1. Right-click on the **Services** folder and choose **Add Class...**. Name it **"UserService"** and click on the **OK** button.
2. In **UserService.cs**, replace everything with the following code piece.  
    ```csharp
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using System.Threading.Tasks;
    using Single_Sign_On_Outlook_Add_inWeb.Models;

    namespace Single_Sign_On_Outlook_Add_inWeb.Services
    {
        public class UserService : IUserService
        {
            private readonly List<UserModel> _users;
            private readonly Dictionary<string, UserModel> _ssoMap;
            private readonly Dictionary<string, byte[]> _saltMap;
            private readonly Random _random;

            public UserService()
            {
                _users = new List<UserModel>();
                _ssoMap = new Dictionary<string, UserModel>();
                _saltMap = new Dictionary<string, byte[]>();
                _random = new Random();

                // Add a new user that we will use in the mail add-in
                _users.Add(new UserModel
                {
                    DisplayName = "Office 365 Developer",
                    Credentials = new CredentialsModel
                    {
                        Username = "#Office365Dev",
                        Password = "#Office365Dev"
                    }
                });
            }

            // Try to get a user service user mapped to an Office 365 user 
            // with the UUID - created by the Exchange Web Services Managed API. 
            public async Task<UserModel> GetUserAsync(string uuid)
            {
                if (_ssoMap.ContainsKey(uuid))
                    return _ssoMap[uuid];
                return null;
            }

            // Get a unique salt for the UUID - which is used when hashing 
            // it for user mapping. If it doesn't exist in the store - it
            // will be randomly generated. 
            public async Task<byte[]> GetUUIDSaltAsync(string uuid)
            {
                // Check if a salt for the UUID already exists
                if (_saltMap.ContainsKey(uuid))
                    return _saltMap[uuid];

                // Generate a new unique salt and add it
                var salt = new byte[16];
                _random.NextBytes(salt);

                _saltMap.Add(uuid, salt);
                return salt;
            }

            // Map a user service user to an Office 365 user with the 
            // UUID. If the mapping is successful, the user is returned.
            public async Task<UserModel> MapUserAsync(string uuid, CredentialsModel credentials)
            {
                var user = _users.FirstOrDefault(u => 
                           u.Credentials.Username.Equals(credentials.Username) &&
                           u.Credentials.Password.Equals(credentials.Password));

                if (user == null)
                    return null;

                _ssoMap.Add(uuid, user);
                return user;
            }
        }
    }

    ```
    
#### Exercise 6.7: Finish the SSO Controller ####
The last step is to complete the Web API Controller. We will create a POST method to handle everything. This method will accept a **UserRequestModel** object to be passed in this request body. The **Credentials** property on the request body is optional, but without it we will not perform a mapping between an Office 365 user and a user in the **UserService** instance. 

Before we do anything that has to do with user lookups or user mappings - we need to validate the identity token in the request. We can of course write this validation code if we want to (<http://simonjaeger.com/dissecting-and-validating-the-exchange-identity-token/>). But we can save ourselves a lot of time (and errors) by using the **Exchange Web Services Managed API** - available as a NuGet package.

When we are mapping an Office 365 user with a user in the **UserService** instance, we don't want to use the Office 365 email address (or Exchange ID for that matter). Instead, we should get the **UniqueUserIdentification** property, salt it and then hash it (this is done in the **GetHashedUUID** method). 
    
1. Click on **Tools** in the Visual Studio 2015 top menu. 
2. Select **NuGet Package Manager** and choose **Package Manager Console**.
   ![](https://raw.githubusercontent.com/simonjaeger/OfficeDev-HOL/master/Single-Sign-On-Outlook-Add-in/Images/PackageManager.png)
3. In the **Package Manager Console**, select the web project (**Single-Sign-On-Outlook-Add-in**) as the **Default project**. 
4. Enter and run **Install-Package Microsoft.Exchange.WebServices** to install the Exchange Web Services Managed API.
   ![](https://raw.githubusercontent.com/simonjaeger/OfficeDev-HOL/master/Single-Sign-On-Outlook-Add-in/Images/PackageManagerConsole.png)
5. In **SSOController.cs** (in the **Controllers** folder), add the following using statements at the top of the file:
    ```csharp
    using Microsoft.Exchange.WebServices.Auth.Validation;
    using Single_Sign_On_Outlook_Add_inWeb.Models;
    using Single_Sign_On_Outlook_Add_inWeb.Services;
    using System;
    using System.Security.Cryptography;
    using System.Text;
    using System.Threading.Tasks;

    ```
2. In **SSOController.cs**, add the following code piece within the class:
    ```csharp
    private static readonly IUserService _userService = new UserService();

    public async Task<UserModel> Post([FromBody]UserRequestModel request)
    {
        try
        {
            // Parse and validate the raw identity token
            var identityToken = (AppIdentityToken)AuthToken.Parse(request.IdentityToken);
            identityToken.Validate(new Uri(request.HostUri));

            // Get the UUID and the salt from the web service
            var uuid = identityToken.UniqueUserIdentification;
            var salt = await _userService.GetUUIDSaltAsync(uuid);

            // Combine and hash it
            var hashedUuid = GetHashedUUID(uuid, salt);

            // Try to find a user mapping with the hashed UUID
            var user = await _userService.GetUserAsync(hashedUuid);

            // If a user mapping is found for the unique user
            // identification - return the user.
            if (user != null)
            {
                return user;
            }

            // End here if there are no credentials in the request
            if (request.Credentials == null)
            {
                return null;
            }

            // Credentials are present, try to map and return a user
            // if the credentials are valid for the web service.
            user = await _userService.MapUserAsync(hashedUuid, request.Credentials);

            if (user == null)
            {
                // Mapping failed.
                return null;
            }
            return user;
        }
        catch (TokenValidationException ex)
        {
            //throw new ApplicationException("An error occurred while validating the identity token.", ex);
            return null;
        }
    }

    private string GetHashedUUID(string uuid, byte[] salt)
    {
        var input = Encoding.UTF8.GetBytes(uuid);
        var saltedInput = new byte[input.Length + salt.Length];

        // Combine the UUID and the salt
        salt.CopyTo(saltedInput, 0);
        input.CopyTo(saltedInput, salt.Length);

        // Compute the hash
        byte[] hash;
        using (var sha265 = new SHA256Managed())
        {
            hash = sha265.ComputeHash(saltedInput);
        }

        return BitConverter.ToString(hash);
    }

    ```
3. Launch your mail add-in and test your work. Use the form to sign in with the credentials defined of a user in the **UserService.cs** class. When the sign in is completed, the mail add-in will reload and automatically sign in. 
   
   If you restart Outlook and the mail add-in, you will find that the user is automatically signed in every time (as long as you keep your application running, because the mapping is in-memory and not persisted).
   ![](https://raw.githubusercontent.com/simonjaeger/OfficeDev-HOL/master/Single-Sign-On-Outlook-Add-in/Images/LaunchedSSOMailAddin7.png)
   
   You could also launch your mail-addin in Outlook Web Access (<https://outlook.office.com/owa/>) and experience the SSO.

# Wrap up  #
View the source code files included in this hands-on lab for a final reference of how your code should be structured (if needed). You should now have grasped an understanding of a few possibilities of interacting with the Office context (a mailbox item in this case). In addition, you have also seen some of the styles and components included in the Office UI Fabric.

# More Resources #
- Discover Office development: <https://msdn.microsoft.com/en-us/office/>
- Learn more about Office UI Fabric: <http://dev.office.com/fabric/>
- Developing Outlook add-ins – where to integrate and what you can do: <http://simonjaeger.com/developing-outlook-add-ins-where-to-integrate-and-what-you-can-do/>
- Single sign-on for Outlook add-ins with ease: <http://simonjaeger.com/single-sign-on-for-outlook-add-ins-with-ease/>