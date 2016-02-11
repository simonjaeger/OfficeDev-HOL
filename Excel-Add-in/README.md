# Hands-on Lab: Excel Add-in #

With the new application model for Office comes a brand new way of extending Office with your own functionality - using the tools and dev stacks that we already know and love. 

This hands-on lab demonstrates a few different ways to interact with the Office context. Adding different types of content, reading selected data from the document and displaying it. In addition, different styles and components from the Office UI Fabric library is used throughout this Office add-in. 
The objective is to get familiar with some of the possiblities that we have when building Excel add-ins.

The hands-on lab is divided into multiple exercises and should be followed in a chronological order. These are the included exercises:

* [1.1 Create the project](#exercise-11-create-the-project)
* [1.2 Edit the manifest](#exercise-12-edit-the-manifest)
* [1.3 Launch the project](#exercise-13-launch-the-project)
* [2.1 Clean up the project](#exercise-21-clean-up-the-project)
* [2.2 Add Office UI Fabric](#exercise-22-add-office-ui-fabric)
* [2.3 Add the base](#exercise-23-add-the-base-css--html)
* [3.1 Add plain text to the document](#exercise-31-add-plain-text-to-the-document)
* [3.2 Add HTML to the document](#exercise-32-add-html-to-the-document)
* [3.3 Add a matrix to the document](#exercise-33-add-a-matrix-to-the-document)
* [3.4 Add an Office Table to the document](#exercise-34-add-an-office-table-to-the-document)
* [3.5 Add Office Open XML to the document](#exercise-35-add-office-open-xml-ooxml-to-the-document)
* [4.1 Add a dialog](#exercise-41-add-a-dialog)
* [4.2 Get selected data as plain text](#exercise-42-get-selected-data-as-plain-text)
* [4.3 Get selected data as HTML](#exercise-43-get-selected-data-as-html)

### Applies to ###
-  Excel Client
-  Excel Online

### Prerequisites ###
- Visual Studio 2015: <https://www.visualstudio.com/en-us/downloads/download-visual-studio-vs.aspx>
- Office Developer Tools: <https://www.visualstudio.com/en-us/features/office-tools-vs.aspx>
- Office 2013 or Office 2016

### Solution ###
Solution | Author(s)
---------|----------
Excel-Add-in | Simon Jäger (**Microsoft**)

### Version history ###
Version  | Date | Comments
---------| -----| --------
1.0  | February 11th 2016 | Initial release

### Disclaimer ###
**THIS CODE IS PROVIDED *AS IS* WITHOUT WARRANTY OF ANY KIND, EITHER EXPRESS OR IMPLIED, INCLUDING ANY IMPLIED WARRANTIES OF FITNESS FOR A PARTICULAR PURPOSE, MERCHANTABILITY, OR NON-INFRINGEMENT.**

----------

# Exercises #

#### Exercise 1.1: Create the project ####
The first thing that we need to do is to create the project itself. Make sure that you have installed all of the required prerequisites before launching Visual Studio 2015. 

1. Click **File**, **New** and finally the **Project** button.
2. In **Templates**, select **Visual C#**, **Office/SharePoint** and then **Office Add-ins**. This will list the Office add-in templates, choose **Office Add-in**. 
3. Name your project **"Excel-Add-in"** and click the **OK** button to continue. 
   ![](https://raw.githubusercontent.com/simonjaeger/OfficeDev-HOL/master/Excel-Add-in/Images/NewProject.png)
4. Next up Visual Studio 2015 will need a bit more information about what you are going to create - in order to set up the required files. Your next step is to decide which type of Office add-in that you want to create. Depending on what you pick, your Office add-in will run in different Office applications and contexts. 
   
   For this hands-on lab, we will create a task pane add-in - this means that our Office add-in will run in a view beside the Office context (e.g. a document, spreadsheet, slide). Select **Task pane** and click on **Next**. 
   ![](https://raw.githubusercontent.com/simonjaeger/OfficeDev-HOL/master/Word-Add-in/Images/TaskpaneAddinType.png)
5. Finally we need to choose the host applications. This means that we are defining the Office applications that our Office (task pane) add-in can run within. Select **Excel** and deselect everything else to create a "Excel-only" add-in. Click **Finish** to complete the wizard.
   ![](https://raw.githubusercontent.com/simonjaeger/OfficeDev-HOL/master/Excel-Add-in/Images/HostApp.png)
6. Using the information you specified in the wizard, Visual Studio 2015 will configure your project. Have a look in the **Solution Explorer** and find your two new projects in the **Excel-Add-In** solution. 
   
   **Excel-Add-in:** This is your manifest project, containing the XML manifest file. This is basically a representation of the information you just specified while creating your Office add-in project. 
   
   **Excel-Add-inWeb:** This is your web project for the Office add-in. This contains all of the different source files that makes up your Excel add-in. We will make quite a few adjustments to this structure as we continue.
   ![](https://raw.githubusercontent.com/simonjaeger/OfficeDev-HOL/master/Excel-Add-in/Images/SolutionExplorer.png)
   
   You've now created the basic structure for a taskpane add-in running in Excel. 

#### Exercise 1.2: Edit the manifest ####
We need to make sure that we understand the manifest file. This file is essential for your add-in; it tells Office where everything is hosted (locally throughout this hands-on lab) and where it can be launched. So let's open it and edit the manifest file.

1. In the manifest project **Excel-Add-in**, double-click the **Excel-Add-inManifest** file. This will open the manifest editor.
   ![](https://raw.githubusercontent.com/simonjaeger/OfficeDev-HOL/master/Excel-Add-in/Images/Manifest.png)
2. In the **General** tab section, find and edit the **Display name** and **Provider name** to anything you'd like.
   ![](https://raw.githubusercontent.com/simonjaeger/OfficeDev-HOL/master/Excel-Add-in/Images/EditManifest.png)
3. Scroll down and pay attention to the **Source location** property. This points to a specific file in your web project (**Excel-Add-inWeb**). When launching your Excel add-in, this page will be the first thing that gets loaded and displayed.

#### Exercise 1.3: Launch the project ####
Before we launch our Excel add-in we should validate that our start actions are proper.

1. Select the manifest project; **Excel-Add-in** in the **Solution Explorer**.                                     
   ![](https://raw.githubusercontent.com/simonjaeger/OfficeDev-HOL/master/Excel-Add-in/Images/SelectManifestProject.png)
2. In the **Properties** window, set **Start Action** to Office Desktop Client. 
3. Set **Start Document** to **[New Excel Document]**.
4. Set **Web Project** to your web project; **Excel-Add-inWeb**.
   ![](https://raw.githubusercontent.com/simonjaeger/OfficeDev-HOL/master/Excel-Add-in/Images/StartActions.png)
5. To launch the project, open on the **Debug** tab at the top of Visual Studio 2015 and click on the **Start Debugging** button. You can also click **Start** in your toolbar or use the **{F5}** keyboard shortcut.            
   ![](https://raw.githubusercontent.com/simonjaeger/OfficeDev-HOL/master/Excel-Add-in/Images/StartProject.png)
6. Once your Excel add-in has launched, you can explore the functionality that comes right out of the box with the Visual Studio 2015 template.            
   ![](https://raw.githubusercontent.com/simonjaeger/OfficeDev-HOL/master/Excel-Add-in/Images/LaunchedAddin.png)


#### Exercise 2.1: Clean up the project ####
While the default styling that comes along with the Visual Studio 2015 template for Office add-ins does its job - leveraging the features of the Office UI Fabric can be fantastic. It's a UI toolkit made specifically for building Office and Office 365 experiences, so it will certainly help us out here.

The Office UI Fabric library comes with everything from styling, components to animations. The majority of the library can be references via a CDN. The heavier parts needs to be downloaded and added to the project itself. We will go through both of these approaches. 

Our first task is to clean up the project, and remove the default styling and setup.

1. Remove the **Content** and **Images** folders from the web project. You can do this by right-clicking these folders in the **Solution Explorer** and choosing the **Delete** option.                                    
    ![](https://raw.githubusercontent.com/simonjaeger/OfficeDev-HOL/master/Word-Add-in/Images/DeleteFolders.png)
2. In your **Solution Explorer**, find the **Home.html** file - which is the startup page for your Excel add-in. **Remove** everything inside the **body** tags.
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
5. In **Home.js**, **remove** the **getDataFromSelection()** function and the call to **app.initialize()**. We are remaking the structure of the Excel add-in, these will no longer be used. You should end up with this:            
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
        <h2 class="ms-font-xxl ms-fontWeight-semibold ms-fontColor-white">HOL: Excel add-in</h2>
    </div>
    <div id="content">
        <!-- Introduction -->
        <p class="ms-font-m ms-fontColor-neutralSecondary">
            This sample demonstrates a few different ways to interact with the Office context.
            Reading and writing data from the current user selection and bindings.
        </p>

        <!-- Exercise Section: Read -->
        <p class="ms-font-l ms-fontWeight-semibold section-title">Exercise: Read</p>
        <p class="ms-font-m ms-fontColor-neutralSecondary">
            Reading data from the user selection is easy, press the button down
            below to test it.
        </p>

        <!-- Exercise: Read data from selection -->
        <div class="section">
        </div>

        <!-- Exercise Section: Write -->
        <p class="ms-font-l ms-fontWeight-semibold section-title">Exercise: Write</p>
        <p class="ms-font-m ms-fontColor-neutralSecondary">
            Writing data to the user selection is also very straight forward,
            press the button down below to test it.
        </p>

        <!-- Exercise: Write data to selection -->
        <div class="section">
        </div>

        <!-- Exercise Section: Bindings -->
        <p class="ms-font-l ms-fontWeight-semibold section-title">Exercise: Bindings</p>
        <p class="ms-font-m ms-fontColor-neutralSecondary">
            With bindings you can read and write data to an area without depending on
            the user selection. Before you can use a binding, you need to create it.
        </p>

        <!-- Exercise: Create bindings -->
        <div class="section">
        </div>

        <!-- Exercise: Write data to binding -->
        <div class="section">
        </div>

        <!-- Exercise: Read data from binding -->
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
3. Launch your Excel add-in to display the new styling. We will add more interactive components in the different sections (inside the recently added HTML piece).                                     
    ![](https://raw.githubusercontent.com/simonjaeger/OfficeDev-HOL/master/Excel-Add-in/Images/LaunchedAddin2.png)
    
#### Exercise 3.1: Read data from selection ####

1. In **Home.html**, locate the "Read data from selection" section (commented) and add the following HTML piece inside the **div** (section) tags. This is an Office UI Fabric styled button. 
    ```html
    <button id="read-data-from-selection" class="ms-Button">
        <span class="ms-Button-label">Read data from selection</span>
    </button>
    
    ```
2. In **Home.js**, add an event handler (below the initialization of the Office UI Fabric components, in the **ready** function) for the click event of the button:
    ```js
    // Add event handlers
    $('#read-data-from-selection').click(readDataFromSelection);
    
    ```
3. In **Home.js**, add the following function to add plain text to the document:
    ```js
    // Read data from the current selection and log it in the JavaScript Console
    function readDataFromSelection() {
        Office.context.document.getSelectedDataAsync(Office.CoercionType.Text,
            function (asyncResult) {
                if (asyncResult.status !== Office.AsyncResultStatus.Succeeded) {
                    // TODO: Handle error
                }
                else {
                    console.log('Selection: ' + asyncResult.value);
                }
            });
    }
    
    ```
4. Launch your Excel add-in and test your work by clicking the **Read data from selection** button. When the button is clicked, the function will be executed; reading the data from the current selection and logging it in the **JavaScript Console**.
5. While your Excel add-in is running, switch back to Visual Studio 2015. Use the **Quick Launch** box in your top right corner to locate the **JavaScript Console** (simply by typing "JavaScript Console").                                     
    ![](https://raw.githubusercontent.com/simonjaeger/OfficeDev-HOL/master/Excel-Add-in/Images/QuickLaunch.png)
6. In the **JavaScript Console**, you should find the selected data logged:                                     
    ![](https://raw.githubusercontent.com/simonjaeger/OfficeDev-HOL/master/Excel-Add-in/Images/SelectedData.png)

#### Exercise 3.2: Write data to selection ####

1. In **Home.html**, locate the "Write data to selection" section (commented) and add the following HTML piece inside the **div** (section) tags. This is an Office UI Fabric styled button. 
    ```html
    <button id="write-data-to-selection" class="ms-Button ms-Button--primary">
        <span class="ms-Button-label">Write data to selection</span>
    </button>
    
    ```
2. In **Home.js**, add an event handler (below the initialization of the Office UI Fabric components, in the **ready** function) for the click event of the button:
    ```js
    $('#write-data-to-selection').click(writeDataToSelection);
    
    ```
3. In **Home.js**, add the following function to add plain text to the document:
    ```js
    // Write data to the current selection 
    function writeDataToSelection() {
        Office.context.document.setSelectedDataAsync('Hello #Office365Dev', {
            coercionType: Office.CoercionType.Text
        }, function (asyncResult) {
            if (asyncResult.status !== Office.AsyncResultStatus.Succeeded) {
                // TODO: Handle error
            }
        });
    }
    
    ```
4. Launch your Excel add-in and test your work by clicking the **Write data to selection** button. When the button is clicked, the function will be executed; writing data to the current selection.                                     
    ![](https://raw.githubusercontent.com/simonjaeger/OfficeDev-HOL/master/Excel-Add-in/Images/Selection.png)

#### Exercise 4.1: Create a custom start document ####
When bulding an Excel add-in, you can use a custom document as the starting point. This is useful when building an Excel add-in that depends on a specific layout or setup. It means that you don't have to set up everything each time you restart your project. 

1. Launch Excel and create a new **Blank workbook**.                                 
    ![](https://raw.githubusercontent.com/simonjaeger/OfficeDev-HOL/master/Excel-Add-in/Images/NewWorkbook.png)
2. In the ribbon menu, go to **Insert** and add a new **Table**. 
3. Select an area which is 4x1 cells in size. Check the box **My table has headers** and click **OK**.                             
    ![](https://raw.githubusercontent.com/simonjaeger/OfficeDev-HOL/master/Excel-Add-in/Images/TableSize.png)
4. Select your new table and go to **Design** in the ribbon menu. 
5. Name your table **"MyTable"** in the **Table Name** field. We will use this name to reference the table in an upcoming exercise.
6. Save and close the Workbook. 

#### Exercise 4.2: Use a custom start document ####
Now let's add the custon start document to our manifest project (**Excel-Add-in**) and configure Visual Studio 2015 to use it whenver we launch the Excel add-in.

1. Right-click your manifest project and choose **Add Existing Item...**.                                      
    ![](https://raw.githubusercontent.com/simonjaeger/OfficeDev-HOL/master/Excel-Add-in/Images/SelectManifestProject.png)
2. Choose your custom start document. 
3. Select the manifest project; **Excel-Add-in** in the **Solution Explorer**.                                      
    ![](https://raw.githubusercontent.com/simonjaeger/OfficeDev-HOL/master/Excel-Add-in/Images/SelectManifestProject.png)
4. Set **Start Document** to your custom start document.                                 
    ![](https://raw.githubusercontent.com/simonjaeger/OfficeDev-HOL/master/Excel-Add-in/Images/CustomStartDocument.png)

The final thing that we need to do is to make sure that our start document launches our Excel add-in when Visual Studio 2015 triggers it. Because this type of action is stored within the document itself, we need to go through a few simple steps to get this going.
1. Launch your project from Visual Studio 2015. You will notice that Excel starts, but your Excel add-in doesn't. 
2. In the ribbon menu, go to **Insert**. 
3. Choose **Add-ins**, **My Add-ins** and finally **Excel-Add-in**.                                  
    ![](https://raw.githubusercontent.com/simonjaeger/OfficeDev-HOL/master/Excel-Add-in/Images/AddAddin.png)
4. Now save your start document, close it and stop debugging. This will save the launch bits of your Excel add-in into the start document.

Launching your project again from Visual Studio 2015 should now trigger the Excel add-in to start with the custom start document.

# Wrap up  #
// TODO

# More Resources #
- Discover Office development: <https://msdn.microsoft.com/en-us/office/>
- Learn more about Office UI Fabric: <http://dev.office.com/fabric/>