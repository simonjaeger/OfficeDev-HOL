# Hands-on Lab: Word Add-in #

### Summary ###

With the new application model for Office comes a brand new way of extending Office with your own functionality - using the tools and dev stacks that we already know and love. 

This hands-on lab demonstrates a few different ways to interact with the Office context. Adding different types of content, reading selected data from the document and displaying it. Additionally - different styles and components from the Office UI Fabric library is used throughout this Office add-in. 
The objective is to get familiar with some of the possiblities that we have when building Word add-ins

The hands-on lab is divided into multiple exercises and should be followed in a chronological order. T

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

#### Exercise 1.0: Creating the project ####
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
   
   You've now created the basic structure for a Word taskpane add-in. 

<a name="#editing-the-manifest"></a>
#### Exercise 1.1: Editing the manifest ####
We need to make sure that we understand the manifest file. This file is essential for your add-in; it tells Office where everything is hosted (locally throughout this hands-on lab) and where it can be launched. So let's open and edit the manifest file.

1. sd
2. sdf
3. sdfsdfds
4. 

1. In the manifest project **Word-Add-in**, double-click the **Word-Add-inManifest** file. This will open the manifest editor.
2. In the **General** tab section, find and edit the **Display name** and **Provider name** to anything you'd like.
3. Scroll down and pay attention to the **Source location** property. This points to a specific file in your web project (**Word-Add-inWeb**). When launching your Word add-in, this page will be the first thing that gets loaded and displayed. 


#### Exercise 1.2: Launching the project ####


#### Exercise 2.1: Creating the project ####



    
    

# More Resources #
- Discover Office development at: <https://msdn.microsoft.com/en-us/office/>
- Get started on Microsoft Azure at: <https://azure.microsoft.com/en-us/>
- Learn about webhooks at: <http://culttt.com/2014/01/22/webhooks/>
- Explore the Outook Notifications REST API and its operations at: <https://msdn.microsoft.com/en-us/office/office365/api/notify-rest-operations> 
- Read more about this sample at: <http://simonjaeger.com/call-me-back-outlook-notifications-rest-api/>