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

The `getAccessToken` API in Office.js enables users who are signed into Office to get access to an AAD-protected add-in and to Microsoft Graph without needing to sign-in again. 

There are two versions of the sample in this repo, each with its own README file:

- In the **Complete** folder is a sample, whose README will walk you through the process of registering the add-in with Azure Active Directory (AAD) and configuring the project.
- In the **SSOAutoSetup** folder is the same sample, but it contains a utility that will automate most of the registration and configuration. Instructions are in the README in that folder. We recommend that you go through the manual process in the Complete folder if you have never registered an app with AAD before. Doing so will give you a better understanding of what AAD does and the significance of the configuration steps.
- **Under construction:** We plan to have a "Before" version which will be used in conjunction with a walkthrough article at [Create a Node.js Office Add-in that uses single sign-on](https://docs.microsoft.com/en-us/office/dev/add-ins/develop/create-sso-office-add-ins-nodejs).

These sample is built on Node.JS, Express, and Microsoft Authentication Library for JavaScript (msal.js). 

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

- Excel on Windows (subscription)
- PowerPoint on Windows (subscription)
- Word on Windows (subscription)

## Prerequisites

To run this code sample, the following are required.

* A code editor. We recommend Visual Studio Code which was used to create the sample.
* An Office 365 account which you can get by joining the [Office 365 Developer Program](https://aka.ms/devprogramsignup) that includes a free 1 year subscription to Office 365. During the preview phase, the SSO requires Office 365 (the subscription version of Office, also called “Click to Run”). You should use the latest monthly version and build from the Insiders channel. You need to be an Office Insider to get this version. For more information, see [Be an Office Insider](https://products.office.com/office-insider?tab=tab-1). 
    > Note: When a build graduates to the production semi-annual channel, support for preview features, including SSO, is turned off for that build.
* At least a few files and folders stored on OneDrive for Business in your Office 365 subscription.
* A Microsoft Azure Tenant. This add-in requires Azure Active Directory (AD). Azure AD provides identity services that applications use for authentication and authorization. A trial subscription can be acquired here: [Microsoft Azure](https://account.windowsazure.com/SignUp).

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

Please go to the README in the **Complete** or **SSOAutoSetup** folder for the next steps.

## Security note

These samples send a hardcoded query parameter on the URL for the Microsoft Graph REST API. If you modify this code in a production add-in and any part of query parameter comes from user input, be sure that it is sanitized so that it cannot be used in a Response header injection attack.

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
