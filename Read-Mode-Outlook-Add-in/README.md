# Hands-on Lab: Read Mode Outlook Add-in #

With the new application model for Office comes a brand new way of extending Office with your own functionality - using the tools and dev stacks that we already know and love. 

This hands-on lab demonstrates a few different ways to interact with the Office context.  Accessing different types of data for a mailbox item (message or appointment) in read mode. In addition, different styles and components from the Office UI Fabric library is used throughout this Office add-in. 
The objective is to get familiar with some of the possiblities that we have when building Excel add-ins.

The hands-on lab is divided into multiple exercises and should be followed in a chronological order. These are the included exercises:

* [1.1 Create the project](#exercise-11-create-the-project)

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
5. Finally we need to choose the supported modes for our Outlook add-in. This means that we are defining the contexts that our Outlook add-in can run within; read, compose or both. If you choose **Read form**, the Outlook add-in will be able to run when a user is viewing a mailbox item. In **Compose form**, the Outlook add-in can run when a user is creating or editing a mailbox item. 
   
   In our case, select **Read form** for both **Email message** and **Appointment**. Deselect everything else to create a read mode Outlook add-in. Click **Finish** to complete the wizard.
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
3. In the **Read Form** tab section, find the **Activation** part. This is what determines the rules for potential activation of your Outlook add-in. By default, both **Item is a message** and **Item is an appointment** should be included. 
4. Scroll down and pay attention to the **Source location** property. This points to a specific file in your web project (**Read-Mode-Outlook-Add-inWeb**). When launching your Word add-in, this page will be the first thing that gets loaded and displayed.       
   ![](https://raw.githubusercontent.com/simonjaeger/OfficeDev-HOL/master/Read-Mode-Outlook-Add-in/Images/ReadFormAddinManifest.png)


# Wrap up  #
View the source code files included in this hands-on lab for a final reference of how your code should be structured (if needed). You should now have grasped an understanding of a few possibilities of interacting with the Office context (a mailbox item in this case). In addition, you have also seen some of the styles and components included in the Office UI Fabric.

# More Resources #
- Discover Office development: <https://msdn.microsoft.com/en-us/office/>
- Learn more about Office UI Fabric: <http://dev.office.com/fabric/>