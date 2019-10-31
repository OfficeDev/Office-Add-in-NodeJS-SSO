---
page_type: sample
products:
- office-excel
- office-powerpoint
- office-word
- office-365
languages:
- javascript
extensions:
  contentType: samples
  technologies:
  - Add-ins
  - Microsoft Graph
  services:
  - Excel
  - Office 365
  createdDate: 5/1/2017 2:09:09 PM
---
# Office Add-in that that supports Single Sign-on to Office, the Add-in, and Microsoft Graph

The `getAccessToken` API in Office.js enables users who are signed into Office to get access to an AAD-protected add-in and to Microsoft Graph without needing to sign-in again. This sample is built on Node.JS, Express, and Microsoft Authentication Library for JavaScript (msal.js). 

 > Note: The `getAccessToken` API is in preview.

## Features

Integrating data from online service providers increases the value and adoption of your add-ins. This code sample shows you how to connect your add-in to Microsoft Graph. Use this code sample to:

* Build an Add-in using Node.js, Express, msal.js, and Office.js. 
* Connect to Microsoft Graph from an Office Add-in.
* Use the OneDrive REST APIs from Microsoft Graph.
* Use the Express routes and middleware to implement the OAuth 2.0 authorization framework in an add-in.
* See how to use the Single Sign-on (SSO) API.
* See how an add-in can fall back to an interactive sign-in in scenarios where SSO is not available.
* Use the msal.js library to implement a fallback authentication/authorization system that is invoked when Office SSO is not available.
* Show a dialog using the Office UI namespace when Office SSO is not available.
* Use add-in commands in an add-in.


## Applies to

-  Excel on Windows (subscription)
-  PowerPoint on Windows (subscription)
-  Word on Windows (subscription)

## Prerequisites

To run this code sample, the following are required.

* A code editor. We recommend Visual Studio Code which was used to create the sample.
* An Office 365 account which you can get by joining the [Office 365 Developer Program](https://aka.ms/devprogramsignup) that includes a free 1 year subscription to Office 365. During the preview phase, the SSO requires Office 365 (the subscription version of Office, also called “Click to Run”). You should use the latest monthly version and build from the Insiders channel. You need to be an Office Insider to get this version. For more information, see [Be an Office Insider](https://products.office.com/office-insider?tab=tab-1). 
    > Note: When a build graduates to the production semi-annual channel, support for preview features, including SSO, is turned off for that build.
* At least a few files and folders stored on OneDrive for Business in your Office 365 subscription.
* A Microsoft Azure Tenant. This add-in requires Azure Active Directiory (AD). Azure AD provides identity services that applications use for authentication and authorization. A trial subscription can be acquired here: [Microsoft Azure](https://account.windowsazure.com/SignUp).

## Solution

Solution | Author(s)
---------|----------
Office Add-in Microsoft Graph ASP.NET | Microsoft

## Version history

Version  | Date | Comments
---------| -----| --------
1.0 | May 10, 2017| Initial release
1.0 | September 15, 2017 | Added support for 2FA.
1.0 | December 8, 2017 | Added extensive error handling.
1.0 | January 7, 2019 | Added information about web application security practices.
2.0 | October 26, 2019 | Changed to use new API and added Display Dialog API fallback.

## Disclaimer

**THIS CODE IS PROVIDED *AS IS* WITHOUT WARRANTY OF ANY KIND, EITHER EXPRESS OR IMPLIED, INCLUDING ANY IMPLIED WARRANTIES OF FITNESS FOR A PARTICULAR PURPOSE, MERCHANTABILITY, OR NON-INFRINGEMENT.**

----------

## To use the project

### Register the add-in
 
1. Register your application in Azure by running the following NPM scipt at the root of your project folder where package.json is located: **npm run configure-sso**

- Your browser will open and prompt for authentication. Enter the user name and passowrd of a user with tenant admin permissions.  If you created an account using [Office 365 Developer Program](https://aka.ms/devprogramsignup), this should suffice.
- Once you have successfully logged in, you will see the scripted steps for registering the application executing in the command shell

### Run the solution

1. Open a command prompt in the root of the project.
2. Run the command `npm start`. 
3. A prompt will appear asking if its OK to register the dev-certifates for the dev-server.  Say 'Yes" to this dialog.  **NOTE:** The dev-certs dialog may not be readily visible if you have many windows open, so you may need to minimize other windows to see it.
4. Excel will automatically start by default.  You can change the default desktop application to Word or PowerPoint by updating the **app-to-debug** in the config section of package.json
5. In the Office application, on the **Home** ribbon, select the **Show Add-in** button in the **SSO Node.js** group to open the task pane add-in.
6. Click the **Get OneDrive File Names** button. If you are logged into Office with either a Work or School (Office 365) account or Microsoft Account, and SSO is working as expected, the first 10 file and folder names in your OneDrive for Business are inserted into the document. (It may take as much as 15 seconds the first time.) If you are not logged in, or you are in a scenario that does not support SSO, or SSO is not working for any reason, you will be prompted to log in. After you log in, the file and folder names appear.

## Security note

The sample sends a hardcoded query parameter on the URL for the Microsoft Graph REST API. If you modify this code in a production add-in and any part of query parameter comes from user input, be sure that it is sanitized so that it cannot be used in a Response header injection attack.

## Questions and comments

We'd love to get your feedback about this sample. You can send your feedback to us in the *Issues* section of this repository.
Questions about developing Office Add-ins should be posted to [Stack Overflow](http://stackoverflow.com). Ensure your questions are tagged with [office-js] and [MicrosoftGraph].


## Additional resources

* [Microsoft Graph documentation](https://docs.microsoft.com/graph/)
* [Office Add-ins documentation](https://docs.microsoft.com/office/dev/add-ins/overview/office-add-ins)

## Copyright
Copyright (c) 2019 Microsoft Corporation. All rights reserved.

This project has adopted the [Microsoft Open Source Code of Conduct](https://opensource.microsoft.com/codeofconduct/). For more information, see the [Code of Conduct FAQ](https://opensource.microsoft.com/codeofconduct/faq/) or contact [opencode@microsoft.com](mailto:opencode@microsoft.com) with any additional questions or comments.

<img src="https://telemetry.sharepointpnp.com/pnp-officeaddins/auth/Office-Add-in-NodeJS-SSO" />


