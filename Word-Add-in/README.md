# Hands-on Lab: Word Add-in #

With the new application model for Office comes a brand new way of extending Office with your own functionality - using the tools and dev stacks that we already know and love. 

This hands-on lab demonstrates a few different ways to interact with the Office context. Adding different types of content, reading selected data from the document and displaying it. In addition, different styles and components from the Office UI Fabric library is used throughout this Office add-in. 
The objective is to get familiar with some of the possiblities that we have when building Word add-ins.

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

Short of time and just want the final sample? Clone this repository and open the solution file: **Word-Add-in\\Source\\Word-Add-in.sln**.           
    ```> git clone https://github.com/simonjaeger/OfficeDev-HOL.git```

### Applies to ###
-  Word Client
-  Word Online

### Prerequisites ###
- Visual Studio 2015: <https://www.visualstudio.com/en-us/downloads/download-visual-studio-vs.aspx>
- Office Developer Tools: <https://www.visualstudio.com/en-us/features/office-tools-vs.aspx>
- Office 2013 (Service Pack 1) or Office 2016

### Solution ###
Solution | Author(s)
---------|----------
Word-Add-in | Simon Jäger (**Microsoft**)

### Version history ###
Version  | Date | Comments
---------| -----| --------
1.1  | February 13th 2016 | Minor updates
1.0  | February 8th 2016 | Initial release

### Disclaimer ###
**THIS CODE IS PROVIDED *AS IS* WITHOUT WARRANTY OF ANY KIND, EITHER EXPRESS OR IMPLIED, INCLUDING ANY IMPLIED WARRANTIES OF FITNESS FOR A PARTICULAR PURPOSE, MERCHANTABILITY, OR NON-INFRINGEMENT.**

----------

# Exercises #

#### Exercise 1.1: Create the project ####
The first thing that we need to do is to create the project itself. Make sure that you have installed all of the required prerequisites before launching Visual Studio 2015. 

1. Click **File**, **New** and finally the **Project** button.
2. In **Templates**, select **Visual C#**, **Office/SharePoint** and then **Office Add-ins**. This will list the Office add-in templates, choose **Office Add-in**. 
3. Name your project **"Word-Add-in"** and click the **OK** button to continue. 
   ![](https://raw.githubusercontent.com/simonjaeger/OfficeDev-HOL/master/Word-Add-in/Images/NewProject.png)
4. Next up Visual Studio 2015 will need a bit more information about what you are going to create - in order to set up the required files. Your next step is to decide which type of Office add-in that you want to create. Depending on what you pick, your Office add-in will run in different Office applications and contexts. 
   
   For this hands-on lab, we will create a task pane add-in - this means that our Office add-in will run in a view beside the Office context (e.g. a document, spreadsheet, slide). Select **Task pane** and click on **Next**. 
   ![](https://raw.githubusercontent.com/simonjaeger/OfficeDev-HOL/master/Word-Add-in/Images/TaskpaneAddinType.png)
5. Finally we need to choose the host applications. This means that we are defining the Office applications that our Office (task pane) add-in can run within. Select **Word** and deselect everything else to create a "Word-only" add-in. Click **Finish** to complete the wizard.
   ![](https://raw.githubusercontent.com/simonjaeger/OfficeDev-HOL/master/Word-Add-in/Images/TaskpaneAddin.png)
6. Using the information you specified in the wizard, Visual Studio 2015 will configure your project. Have a look in the **Solution Explorer** and find your two new projects in the **Word-Add-In** solution. 
   
   **Word-Add-in:** This is your manifest project, containing the XML manifest file. This is basically a representation of the information you just specified while creating your Office add-in project. 
   
   **Word-Add-inWeb:** This is your web project for the Office add-in. This contains all of the different source files that makes up your Word add-in. We will make quite a few adjustments to this structure as we continue.
   ![](https://raw.githubusercontent.com/simonjaeger/OfficeDev-HOL/master/Word-Add-in/Images/SolutionExplorer.png)
   
   You've now created the basic structure for a taskpane add-in running in Word. 

#### Exercise 1.2: Edit the manifest ####
We need to make sure that we understand the manifest file. This file is essential for your add-in; it tells Office where everything is hosted (locally throughout this hands-on lab) and where it can be launched. So let's open it and edit the manifest file.

1. In the manifest project **Word-Add-in**, double-click the **Word-Add-inManifest** file. This will open the manifest editor.
   ![](https://raw.githubusercontent.com/simonjaeger/OfficeDev-HOL/master/Word-Add-in/Images/WordAddinManifest.png)
2. In the **General** tab section, find and edit the **Display name** and **Provider name** to anything you'd like.
3. Scroll down and pay attention to the **Source location** property. This points to a specific file in your web project (**Word-Add-inWeb**). When launching your Word add-in, this page will be the first thing that gets loaded and displayed. 
   ![](https://raw.githubusercontent.com/simonjaeger/OfficeDev-HOL/master/Word-Add-in/Images/EditManifest.png)

#### Exercise 1.3: Launch the project ####
Before we launch our Word add-in we should validate that our start actions are proper.

1. Select the manifest project; **Word-Add-in** in the **Solution Explorer**.                                     
   ![](https://raw.githubusercontent.com/simonjaeger/OfficeDev-HOL/master/Word-Add-in/Images/SelectManifestProject.png)
2. In the **Properties** window, set **Start Action** to Office Desktop Client. 
3. Set **Start Document** to **[New Word Document]**.
4. Set **Web Project** to your web project; **Word-Add-inWeb**.
   ![](https://raw.githubusercontent.com/simonjaeger/OfficeDev-HOL/master/Word-Add-in/Images/StartActions.png)
5. To launch the project, open on the **Debug** tab at the top of Visual Studio 2015 and click on the **Start Debugging** button. You can also click **Start** in your toolbar or use the **{F5}** keyboard shortcut.            
   ![](https://raw.githubusercontent.com/simonjaeger/OfficeDev-HOL/master/Word-Add-in/Images/StartProject.png)
6. Once your Word add-in has launched, you can explore the functionality that comes right out of the box with the Visual Studio 2015 template.            
   ![](https://raw.githubusercontent.com/simonjaeger/OfficeDev-HOL/master/Word-Add-in/Images/LaunchedAddin.png)


#### Exercise 2.1: Clean up the project ####
While the default styling that comes along with the Visual Studio 2015 template for Office add-ins does its job - leveraging the features of the Office UI Fabric can be fantastic. It's a UI toolkit made specifically for building Office and Office 365 experiences, so it will certainly help us out here.

The Office UI Fabric library comes with everything from styling, components to animations. The majority of the library can be references via a CDN. The heavier parts needs to be downloaded and added to the project itself. We will go through both of these approaches. 

Our first task is to clean up the project, and remove the default styling and setup.

1. Remove the **Content** and **Images** folders from the web project. You can do this by right-clicking these folders in the **Solution Explorer** and choosing the **Delete** option.                                    
    ![](https://raw.githubusercontent.com/simonjaeger/OfficeDev-HOL/master/Word-Add-in/Images/DeleteFolders.png)
2. In your **Solution Explorer**, find the **Home.html** file - which is the startup page for your Word add-in. **Remove** everything inside the **body** tags. 
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
5. In **Home.js**, **remove** the **getDataFromSelection()** function and the call to **app.initialize()**. We are remaking the structure of the Word add-in, these will no longer be used. You should end up with this:            
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
2. Some components in the Office UI Fabric library require some additional JavaScript to function. In our case, we will use a Dropdown component that needs this. **Download** the JavaScript file for this component (**Jquery.Dropdown.js**) at <https://raw.githubusercontent.com/OfficeDev/Office-UI-Fabric/master/src/components/Dropdown/Jquery.Dropdown.js> or get it by browsing the files included in this hands-on lab. 
3. Add the **Jquery.Dropdown.js** file to your **Scripts** folder in the **Solution Explorer**. You can do this by right-clicking the **Scripts** folder and choosing **Add Existing Item...**.                                    
    ![](https://raw.githubusercontent.com/simonjaeger/OfficeDev-HOL/master/Word-Add-in/Images/AddExisting.png)
4. In **Home.html**, reference the **Jquery.Dropdown.js** file by adding the following line inside the **head** tags. Be sure to add it after the reference to **"../../Scripts/jquery-1.9.1.js"**.                      
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
        <h2 class="ms-font-xxl ms-fontWeight-semibold ms-fontColor-white">HOL: Word add-in</h2>
    </div>
    <div id="content">
        <!-- Introduction -->
        <p class="ms-font-m ms-fontColor-neutralSecondary">
            This sample demonstrates a few different ways to interact with the Office context. Adding different types of content and reading data from the document.
        </p>

        <!-- Exercise Section: Write -->
        <p class="ms-font-l ms-fontWeight-semibold section-title">Exercise: Write</p>
        <p class="ms-font-m ms-fontColor-neutralSecondary">
            Here is a few different ways for adding content, using different data types.
        </p>

        <!-- Exercise: Add plain text and HTML -->
        <div class="section">
        </div>

        <!-- Exercise: Add matrix -->
        <div class="section">
            <!-- TODO: Replace with code -->
        </div>

        <!-- Exercise: Add Office Table -->
        <div class="section">
        </div>

        <!-- Exercise: Add OOXML -->
        <div class="section">
        </div>

        <!-- Exercise Section: Read -->
        <p class="ms-font-l ms-fontWeight-semibold section-title">Exercise: Read</p>
        <p class="ms-font-m ms-fontColor-neutralSecondary">
            Here is a couple of different functions for getting the selected data, in two different formats.
        </p>

        <!-- Exercise: Selected data dialog -->


        <!-- Exercise: Get selected data (plain text) -->
        <div class="section">
        </div>

        <!-- Exercise: Get selected data (HTML) -->
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
3. Launch your Word add-in to display the new styling. We will add more interactive components in the different sections (inside the recently added HTML piece).                                     
    ![](https://raw.githubusercontent.com/simonjaeger/OfficeDev-HOL/master/Word-Add-in/Images/LaunchedAddin2.png)
    
#### Exercise 3.1: Add plain text to the document ####

1. In **Home.html**, locate the "Exercise: Add plain text and HTML" section (commented) and add the following HTML piece inside the **div** (section) tags. This is an Office UI Fabric styled button. 
    ```html
    <button id="add-plain-text" class="ms-Button">
        <span class="ms-Button-label">Add plain text</span>
    </button>
    
    ```
2. In **Home.js**, add an event handler (below the initialization of the Office UI Fabric components, in the **ready** function) for the click event of the button:
    ```js
    // Add event handlers
    $('#add-plain-text').click(addPlainText);
    
    ```
3. In **Home.js**, add the following function to add plain text to the document:
    ```js
    // Adds data (plain text) to the current document selection
    function addPlainText() {
        var text = 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod' +
            'tempor incididunt ut labore et dolore magna aliqua. Ut enim ad minim veniam, ' +
            'quis nostrud exercitation ullamco laboris nisi ut aliquip ex ea commodo consequat. ' +
            'Duis aute irure dolor in reprehenderit in voluptate velit esse cillum dolore ' +
            'eu fugiat nulla pariatur. Excepteur sint occaecat cupidatat non proident, sunt ' +
            'in culpa qui officia deserunt mollit anim id est laborum.';

        // Set selection
        Office.context.document.setSelectedDataAsync(text, {
            coercionType: Office.CoercionType.Text
        }, onSelectionSet);

    }
    ```
4. In **Home.js**, add the following function to serve as a callback when adding any data to the selection in the document. You can perform validation checks in this function and present errors if something goes wrong during the insertion.
    ```js
    // Callback function for the asynchronous write function
    function onSelectionSet(asyncResult) {
        if (asyncResult.status !== Office.AsyncResultStatus.Succeeded) {
            // TODO: Handle error
        }
    }
    
    ```
5. Launch your Word add-in and test your work by clicking the **Add plain text** button. When the button is clicked, the function will be executed; adding a piece of plain text into the document.

#### Exercise 3.2: Add HTML to the document ####

1. In **Home.html**, locate the "Exercise: Add plain text and HTML" section (commented) and add the following HTML piece inside the **div** (section) tags. This is an Office UI Fabric styled button. 
    ```html
    <button id="add-html" class="ms-Button ms-Button--primary">
        <span class="ms-Button-label">Add HTML</span>
    </button>
    
    ```
2. In **Home.js**, add an event handler (below the initialization of the Office UI Fabric components, in the **ready** function) for the click event of the button:
    ```js
    $('#add-html').click(addHtml);
    
    ```
3. In **Home.js**, add the following function to add HTML to the document:
    ```js
    // Add data (HTML) to the current document selection
    function addHtml() {
        var elements = $('<div>')
            .append($('<h2>').text('Lorem ipsum dolor'))
            .append($('<p>').html('Duis aute irure dolor in <strong>reprehenderit in</strong>' +
                                  'voluptate velit esse cillum dolore.'));
        var html = elements.html();

        // Set selection
        Office.context.document.setSelectedDataAsync(html, {
            coercionType: Office.CoercionType.Html
        }, onSelectionSet);
    }
    ```
4. Launch your Word add-in and test your work by clicking the **Add HTML** button. When the button is clicked, the function will be executed; adding a piece of HTML into the document.

#### Exercise 3.3: Add a matrix to the document ####

1. In **Home.html**, locate the "Exercise: Add matrix" section (commented) and add the following HTML piece inside the **div** (section) tags. This is an Office UI Fabric styled button. 
    ```html
    <button id="add-matrix" class="ms-Button ms-Button--compound">
        <span class="ms-Button-label">Add matrix</span>
        <span class="ms-Button-description">
            Description of the action this button takes
        </span>
    </button>
    
    ```
2. In **Home.js**, add an event handler (below the initialization of the Office UI Fabric components, in the **ready** function) for the click event of the button:
    ```js
    $('#add-matrix').click(addMatrix);
    
    ```
3. In **Home.js**, add the following function to add a matrix to the document:
    ```js
    // Add data (matrix) to the current document selection
    function addMatrix() {
        var matrix = [["Header", "Header"],
                ["Entry", "Entry"],
                ["Entry", "Entry"],
                ["Entry", "Entry"]];

        // Set selection
        Office.context.document.setSelectedDataAsync(matrix, {
            coercionType: Office.CoercionType.Matrix
        }, onSelectionSet);
    }
    ```
4. Launch your Word add-in and test your work by clicking the **Add matrix** button. When the button is clicked, the function will be executed; adding a matrix as a table into the document.

#### Exercise 3.4: Add an Office Table to the document ####

1. In **Home.html**, locate the "Add Office Table" section (commented) and add the following HTML piece inside the **div** (section) tags. This is an Office UI Fabric styled button. 
    ```html
    <button id="add-office-table" class="ms-Button ms-Button--command">
        <span class="ms-Button-icon">
            <i class="ms-Icon ms-Icon--plus"></i>
        </span>
        <span class="ms-Button-label">Add Office Table</span>
    </button>
    
    ```
2. In **Home.js**, add an event handler (below the initialization of the Office UI Fabric components, in the **ready** function) for the click event of the button:
    ```js
    $('#add-office-table').click(addOfficeTable);
    
    ```
3. In **Home.js**, add the following function to add an Office Table object to the document:
    ```js
    // Add data (Office Table) to the current document selection
    function addOfficeTable() {
        var table = new Office.TableData();
        table.headers = [['Header', 'Header']];
        table.rows = [['Entry', 'Entry'], ['Entry', 'Entry'], ['Entry', 'Entry']];

        // Set selection
        Office.context.document.setSelectedDataAsync(table, {
            coercionType: Office.CoercionType.Table
        }, onSelectionSet);
    }
    ```
4. Launch your Word add-in and test your work by clicking the **Add Office Table** button. When the button is clicked, the function will be executed; adding an Office Table object as a table into the document.

#### Exercise 3.5: Add Office Open XML (OOXML) to the document ####

1. In **Home.html**, locate the "Exercise: Add OOXML" section (commented) and add the following HTML piece inside the **div** (section) tags. This is an Office UI Fabric styled Dropdown (using the **Jquery.Dropdown.js** file) and button. 
    ```html
    <div class="ms-Dropdown" tabindex="0">
        <i class="ms-Dropdown-caretDown ms-Icon ms-Icon--caretDown"></i>
        <select id="ooxml-file" class="ms-Dropdown-select">
            <option value="Chart.xml">Chart.xml</option>
            <option value="SimpleImage.xml">SimpleImage.xml</option>
            <option value="SmartArt.xml">SmartArt.xml</option>
            <option value="TableStyled.xml">TableStyled.xml</option>
            <option value="TableWithDirectFormat.xml">TableWithDirectFormat.xml</option>
            <option value="TextBoxWordArt.xml">TextBoxWordArt.xml</option>
            <option value="TextWithStyle.xml">TextWithStyle.xml</option>
        </select>
    </div>

    <button id="add-open-xml" class="ms-Button ms-Button--hero">
        <span class="ms-Button-icon">
            <i class="ms-Icon ms-Icon--plus">
            </i>
        </span>
        <span class="ms-Button-label">Add OOXML</span>
    </button>
    
    ```
2. In **Home.js**, add an event handler (below the initialization of the Office UI Fabric components, in the **ready** function) for the click event of the button:
    ```js
    $('#add-open-xml').click(addOpenXml);
    
    ```
3. In **Home.js**, add the following function to add OOXML (read from the selected file) to the document:
    ```js
    // Add data (OOXM) to the current document selection
    function addOpenXml() {
        // Get the selected file
        var file = $('#ooxml-file').val();

        // Get the file contents
        $.ajax({
            url: '/../../OOXML/' + file,
            type: 'GET',
            dataType: 'text',
            success: function (data) {
                // Set selection
                Office.context.document.setSelectedDataAsync(data, {
                    coercionType: Office.CoercionType.Ooxml
                }, onSelectionSet);
            },
            error: function (e) {
                // TODO: Handle error
            }
        });
    }
    ```
4. Browse the files included in this hands-on lab or head over to <https://github.com/OfficeDev/Word-Add-in-Load-and-write-Open-XML/tree/master/C%23/LoadingAndWritingOOXMLWeb/OOXMLSamples>. Get ahold of the listed OOXML files. You can open these files in any text editor and explore the OOXML data.
      * Chart.xml
      * SimpleImage.xml
      * TableStyled.xml
      * TableWithDirectFormat.xml
      * TextBoxWordArt.xml
      * TextWithStyle.xml
5. Create a new folder in your web project and name it **OOXML** in the **Solution Explorer**. Add these files into this folder by right-clicking it and choosing **Add Existing Item...**.                                      
    ![](https://raw.githubusercontent.com/simonjaeger/OfficeDev-HOL/master/Word-Add-in/Images/OOXML.png)
6. Launch your Word add-in and test your work by clicking the **Add OOXML** button. When the button is clicked, the function will be executed; adding an OOXML piece (read from the selected file) into the document.                                   
    ![](https://raw.githubusercontent.com/simonjaeger/OfficeDev-HOL/master/Word-Add-in/Images/Chart.png) 

#### Exercise 4.1: Add a dialog ####

1. In **Home.html**, locate the "Exercise: Selected data dialog" section (commented) and add the following HTML piece right after. This is an Office UI Fabric styled dialog. 
    ```html
    <div id="selected-data-dialog" class="ms-Dialog ms-Dialog--lgHeader">
        <div class="ms-Overlay"></div>
        <div class="ms-Dialog-main">
            <button class="ms-Dialog-button ms-Dialog-button--close"><i class="ms-Icon ms-Icon--x"></i></button>
            <div class="ms-Dialog-header">
                <p class="ms-Dialog-title">
                    Selected data
                </p>
            </div>
            <div class="ms-Dialog-inner">
                <div class="ms-Dialog-content">
                    <p id="selected-data-dialog-text" class="ms-Dialog-subText" style="max-height: 95px; overflow-y: hidden;">
                        <!-- Dialog text -->
                    </p>
                </div>
                <div class="ms-Dialog-actions">
                    <div class="ms-Dialog-actionsRight">
                        <button id="selected-data-dialog-more" class="ms-Dialog-action ms-Button ms-Button--primary">
                            <span class="ms-Button-label">OK</span>
                        </button>
                        <button id="selected-data-dialog-got-it" class="ms-Dialog-action ms-Button">
                            <span class="ms-Button-label">Got it</span>
                        </button>
                    </div>
                </div>
            </div>
        </div>
    </div>
    
    ```
2. Launch your Word add-in and test your work. You should find a dialog covering up most of your display area.
    ![](https://raw.githubusercontent.com/simonjaeger/OfficeDev-HOL/master/Word-Add-in/Images/Dialog.png)
3. In **Home.js**, add two event handlers (below the initialization of the Office UI Fabric components, in the **ready** function) for the dialog. This will allow us to close it.
    ```js
    $('#selected-data-dialog-more').click(hideSelectedDataDialog);
    $('#selected-data-dialog-got-it').click(hideSelectedDataDialog);
    
    ```
4. In **Home.js**, add the following functions to close the dialog:
    ```js
    // Show the selected data dialog
    function showSelectedDataDialog(text) {
        if (text.length === 0) {
            text = 'No selected data was found';
        }

        // Set the dialog text
        $('#selected-data-dialog-text').text(text);
        $('#selected-data-dialog').show();
    }

    // Hide the selected data dialog
    function hideSelectedDataDialog() {
        $('#selected-data-dialog').hide();
    }
    ```
5. Launch your Word add-in and test your work. You should be able to close the dialog using any of the two buttons.
6. In **Home.css**, add the following CSS piece to hide the dialog when your Word add-in has launched.
    ```css
    #selected-data-dialog {
        display: none;
    }
    
    ```

#### Exercise 4.2: Get selected data as plain text  ####

1. In **Home.html**, locate the "Exercise: Get selected data (plain text)" section (commented) and add the following HTML piece inside the **div** (section) tags. This is an Office UI Fabric styled button. 
    ```html
    <button id="get-selected-plain-text" class="ms-Button">
        <span class="ms-Button-label">Get selected data (plain text)</span>
    </button>
    
    ```
2. In **Home.js**, add an event handler (below the initialization of the Office UI Fabric components, in the **ready** function) for the click event of the button:
    ```js
    $('#get-selected-plain-text').click(getSelectedPlainText);
    
    ```
3. In **Home.js**, add the following functions to get the selected data as plain text and then display it:
    ```js
    // Get the selected data as plain text
    function getSelectedPlainText() {
        getSelectedData(Office.CoercionType.Text);
    }

    // Get the selected data
    function getSelectedData(coercionType) {
        Office.context.document.getSelectedDataAsync(coercionType, function (asyncResult) {
            if (asyncResult.status !== Office.AsyncResultStatus.Succeeded) {
                // TODO: Handle error
            }
            else {
                showSelectedDataDialog(asyncResult.value);
            }
        });
    }
    ```
4. Launch your Word add-in and test your work by clicking the **Get selected data (plain text)** button. When the button is clicked, the function will be executed; getting the selected data as plain text and displaying it using the dialog.

#### Exercise 4.3: Get selected data as HTML  ####

1. In **Home.html**, locate the "Exercise: Get selected data (HTML)" section (commented) and add the following HTML piece inside the **div** (section) tags. This is an Office UI Fabric styled button. 
    ```html
    <button id="get-selected-html" class="ms-Button ms-Button--primary">
        <span class="ms-Button-label">Get selected data (HTML)</span>
    </button>
    
    ```
2. In **Home.js**, add an event handler (below the initialization of the Office UI Fabric components, in the **ready** function) for the click event of the button:
    ```js
    $('#get-selected-html').click(getSelectedHTML);
    
    ```
3. In **Home.js**, add the following functions to get the selected data as plain text and then display it:
    ```js
    // Get the selected data as HTML
    function getSelectedHTML() {
        getSelectedData(Office.CoercionType.Html);
    }
    
    ```
4. Launch your Word add-in and test your work by clicking the **Get selected data (HTML)** button. When the button is clicked, the function will be executed; getting the selected data as HTML and displaying it using the dialog.                                   
    ![](https://raw.githubusercontent.com/simonjaeger/OfficeDev-HOL/master/Word-Add-in/Images/GetSelectedData.png) 

# Wrap up  #
View the source code files included in this hands-on lab for a final reference of how your code should be structured (if needed). You should now have grasped an understanding of a few possibilities of interacting with the Office context (a document in this case). In addition, you have also seen some of the styles and components included in the Office UI Fabric. 

# More Resources #
- Discover Office development: <https://msdn.microsoft.com/en-us/office/>
- Learn more about Office UI Fabric: <http://dev.office.com/fabric/>