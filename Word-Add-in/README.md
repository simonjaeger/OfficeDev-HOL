# Hands-on Lab: Word Add-in #

### Summary ###

With the new application model for Office comes a brand new way of extending Office with your own functionality - using the tools and dev stacks that we already know and love. 

This hands-on lab demonstrates a few different ways to interact with the Office context. Adding different types of content, reading selected data from the document and displaying it. Additionally - different styles and components from the Office UI Fabric library is used throughout this Office add-in. 
The objective is to get familiar with some of the possiblities that we have when building Word add-ins

The hands-on lab is divided into multiple exercises and should be followed in a chronological order. These are the included exercises:

* [1.0 Creating the project](#exercise-10-creating-the-project)
* [1.1 Editing the manifest](#exercise-11-editing-the-manifest)
* [1.2 Launching the project](#exercise-12-launching-the-project)

### Applies to ###
-  Word Client
-  Word Online

### Prerequisites ###
// TODO

### Solution ###
Solution | Author(s)
---------|----------
Word-Add-in | Simon Jäger (**Microsoft**)

### Version history ###
Version  | Date | Comments
---------| -----| --------
1.0  | February 8th 2016 | Initial release

### Disclaimer ###
**THIS CODE IS PROVIDED *AS IS* WITHOUT WARRANTY OF ANY KIND, EITHER EXPRESS OR IMPLIED, INCLUDING ANY IMPLIED WARRANTIES OF FITNESS FOR A PARTICULAR PURPOSE, MERCHANTABILITY, OR NON-INFRINGEMENT.**

----------

# Exercises #

#### Exercise 1.0: Create the project ####
The first thing that we need to do is to create the project itself. Make sure that you have installed all of the required prerequisites and Launch Visual Studio 2015. 

1. Click **File**, **New** and finally the **Project** button.
2. In **Templates**, select **Visual C#**, **Office/SharePoint** and then **Office Add-ins**. This will list the Office add-in templates, choose **Office Add-in**. 
3. Name your project **"Word-Add-in"** and click the **OK** button to continue. 
   ![](https://raw.githubusercontent.com/simonjaeger/OfficeDev-HOL/master/Word-Add-in/Images/NewProject.png)
4. Next up Visual Studio 2015 will need a bit more information about what you are going to create - in order to set up the required files. Your next step is to decide which type of Office add-in that you want to create. Depending on what you pick, your Office add-in will run in different places and Office applications. 
   
   For this hands-on lab, we will create a task pane add-in - this means that our Office add-in will run in task-pane beside the Office context (e.g. a document, spreadsheet, slide). Select **Task pane** and click on **Next**. 
   ![](https://raw.githubusercontent.com/simonjaeger/OfficeDev-HOL/master/Word-Add-in/Images/AddinType.png)
5. Finally we need to choose the application. This means that we are defining the Office applications that our Office (taskpane) add-in can run within. Select **Word** and deselect everything else to create a Word add-in. Click **Finish** to complete the wizard.
   ![](https://raw.githubusercontent.com/simonjaeger/OfficeDev-HOL/master/Word-Add-in/Images/TaskpaneAddin.png)
6. Using the information you specified in the wizard, Visual Studio 2015 will configure your project. Have a look in the **Solution Explorer** and find your two new projects in the **Word-Add-In** solution. 
   
   **Word-Add-in:** This is your manifest project, containing the XML manifest file. This is basically a representation of the information you just specified while creating your Office add-in project. 
   
   **Word-Add-inWeb:** This is your web project for the Office add-in. This contains all of the different source files that makes up your Word add-in. We will make quite a few adjustments to this structure as we continue.
   ![](https://raw.githubusercontent.com/simonjaeger/OfficeDev-HOL/master/Word-Add-in/Images/SolutionExplorer.png)
   
   You've now created the basic structure for a taskpane add-in running in Word. 

#### Exercise 1.1: Edit the manifest ####
We need to make sure that we understand the manifest file. This file is essential for your add-in; it tells Office where everything is hosted (locally throughout this hands-on lab) and where it can be launched. So let's open and edit the manifest file.

1. In the manifest project **Word-Add-in**, double-click the **Word-Add-inManifest** file. This will open the manifest editor.
   ![](https://raw.githubusercontent.com/simonjaeger/OfficeDev-HOL/master/Word-Add-in/Images/WordAddinManifest.png)
2. In the **General** tab section, find and edit the **Display name** and **Provider name** to anything you'd like.
   ![](https://raw.githubusercontent.com/simonjaeger/OfficeDev-HOL/master/Word-Add-in/Images/EditManifest.png)
3. Scroll down and pay attention to the **Source location** property. This points to a specific file in your web project (**Word-Add-inWeb**). When launching your Word add-in, this page will be the first thing that gets loaded and displayed.

#### Exercise 1.2: Launch the project ####
Before we launch our Word add-in we should validate that our start actions are proper.

1. Select the manifest project; **Word-Add-in** in the **Solution Explorer**.                                     
   ![](https://raw.githubusercontent.com/simonjaeger/OfficeDev-HOL/master/Word-Add-in/Images/SelectManifestProject.png)
2. In the **Properties** window, set **Start Action** to Office Desktop Client. 
3. Set **Start Document** to **[New Word Document]**.
4. Set **Web Project** to your web project; **Word-Add-inWeb**.
   ![](https://raw.githubusercontent.com/simonjaeger/OfficeDev-HOL/master/Word-Add-in/Images/StartActions.png)
5. To launch the project, open on the **Debug** tab in the top and click on the **Start Debugging** button. You can also click **Start** in your toolbar or use the **F5** keyboard shortcut.            
   ![](https://raw.githubusercontent.com/simonjaeger/OfficeDev-HOL/master/Word-Add-in/Images/StartProject.png)
6. Once your Word add-in has launched, you can explore the functionality that comes right of the box with the Visual Studio 2015 template.            
   ![](https://raw.githubusercontent.com/simonjaeger/OfficeDev-HOL/master/Word-Add-in/Images/LaunchedAddin.png)


#### Exercise 2.0: Clean up the project ####
While the default styling that comes along with the Visual Studio 2015 template for Office add-ins does its job - leveraging the features of the Office UI Fabric can be fantastic. It's a UI toolit made specifically for building Office and Office 365 experiences, so it will certainly help us here.

The Office UI Fabric library comes with everything from styling, components to animations. The majority of the library can be references via a CDN. The heavier parts needs to be downloaded and added to the project itself. We will go through both of these things. 

Our first task here is to clean up the project.

1. Remove the **Content** and **Images** folders from the project. You can do this by right-clicking these folders in the **Solution Explorer** and choosing the **Delete** option.                                    
    ![](https://raw.githubusercontent.com/simonjaeger/OfficeDev-HOL/master/Word-Add-in/Images/DeleteFolders.png)
2. In your **Solution Explorer**, find the **Home.html** file - which is the startup file for your Word add-in. **Remove** everything in the **<body>** tag. This should leave you with something like this:
    ```html
    <!DOCTYPE html>
    <html>
    <head>
        <meta charset="UTF-8" />
        <meta http-equiv="X-UA-Compatible" content="IE=Edge" />
        <title></title>
        <script src="../../Scripts/jquery-1.9.1.js" type="text/javascript"></script>
    
        <link href="../../Content/Office.css" rel="stylesheet" type="text/css" />
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
3. In **App.js**, **remove** the **initialize()** function defined on the **app** object, as this will not be used:            
    ```js
    var app = (function () {
        "use strict";
    
        var app = {};
        return app;
    })();
    ```
4. In **Home.js**, **remove** the **getDataFromSelection()** function and the call to **app.initialize()**. We are remaking the structure of the Word add-in, these will no longer be used. You should end up with this:            
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
5. In **App.css**, **remove** everything, leaving you with an empty file.


#### Exercise 2.1: Add Office UI Fabric ####
1. In **Home.html**, add two CSS references to the CDN for Office UI Fabric in the **<head>** tag. Add them before the CSS reference to **"../App.css"**.
    ```xml
    <link rel="stylesheet" href="https://appsforoffice.microsoft.com/fabric/1.0/fabric.min.css">
    <link rel="stylesheet" href="https://appsforoffice.microsoft.com/fabric/1.0/fabric.components.min.css">
    
    ```
2. Some components in the Office UI Fabric library require some additional JavaScript to function. In our case, we will use a Dropdown component that needs this. **Download** the JavaScript file (**Jquery.Dropdown.js**) at <https://raw.githubusercontent.com/OfficeDev/Office-UI-Fabric/master/src/components/Dropdown/Jquery.Dropdown.js> or get it by browsing the files included in this hands-on lab. 
3. Add the **Jquery.Dropdown.js** file to your **Scripts** folder in the **Solution Explorer**. You can do this by right-clicking the **Scripts** folder and choose **Add Existing Item**.                                    
    ![](https://raw.githubusercontent.com/simonjaeger/OfficeDev-HOL/master/Word-Add-in/Images/AddExisting.png)
4. In **Home.html**, reference the **Jquery.Dropdown.js** file by adding the following line in the **<head>** tag. Be sure to add it after the reference to **"../../Scripts/jquery-1.9.1.js"**.                      
    ```html
    <script src="../../Scripts/Jquery.Dropdown.js" type="text/javascript"></script>
    ```
5. In **Home.js**, add the following line in the **ready** function of your page.             
    ```js
    $(".ms-Dropdown").Dropdown();
    ```  
    This will use the functionality within the **Jquery.Dropdown.js** file and initialize any Dropdown components that we add to the view. 
    
6. Your **Home.html** file should now look like this: 
    ```html
    <!DOCTYPE html>
    <html>
    <head>
        <meta charset="UTF-8" />
        <meta http-equiv="X-UA-Compatible" content="IE=Edge" />
        <title></title>
        <script src="../../Scripts/jquery-1.9.1.js" type="text/javascript"></script>

        <link href="../../Content/Office.css" rel="stylesheet" type="text/css" />
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
    Your **Home.js** file should look like this:
    ```js
    (function () {
        "use strict";
    
        // The initialize function must be run each time a new page is loaded
        Office.initialize = function (reason) {
           $(document).ready(function () {
               // Initialize Office UI Fabric components (dropdowns)
               $(".ms-Dropdown").Dropdown();
            });
        };
    })();
    ``` 
6. 
    
    

# More Resources #
- Discover Office development at: <https://msdn.microsoft.com/en-us/office/>
- Office UI Fabric