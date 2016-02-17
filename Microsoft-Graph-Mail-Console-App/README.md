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
#### Exercise 1.1: Create the project ####
The first thing that we need to do is to create the project itself. Make sure that you have installed all of the required prerequisites before launching Visual Studio 2015. 

1. Click **File**, **New** and finally the **Project** button.
2. In **Templates**, select **Visual C#**, **Windows** and then choose **Console Application**. 
3. Name your project **"Microsoft-Graph-Mail-Console-App"** and click the **OK** button to continue. 
   ![](https://raw.githubusercontent.com/simonjaeger/OfficeDev-HOL/master/Microsoft-Graph-Mail-Console-App/Images/NewProject.png)




# Wrap up  #
View the source code files included in this hands-on lab for a final reference of how your code should be structured (if needed). You should now have grasped an understanding of a few possibilities of interacting (and authenticating) with the Microsoft Graph.

# More Resources #
- Discover Office development: <https://msdn.microsoft.com/en-us/office/>
- Learn more about the Microsoft Graph: <http://graph.microsoft.io/en-us/>
- Microsoft Graph: Authentication with Azure AD: <http://simonjaeger.com/microsoft-graph-authentication-with-azure-ad/>
- Microsoft Graph: Authentication with the converged model (preview): <http://simonjaeger.com/microsoft-graph-authentication-with-the-converged-model-preview/>