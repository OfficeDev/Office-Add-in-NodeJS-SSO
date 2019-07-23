---
page_type: sample
products:
- office-excel
- office-word
- office-powerpoint
- office-project
- office-outlook
- office-365
languages:
- javascript
- typescript
extensions:
  contentType: samples
  technologies:
  - Add-ins
  - Microsoft Graph
  services:
  - Excel
  - Outlook
  - Office 365
  createdDate: 5/3/2017 2:24:40 PM
---
# Office Add-in that supports Single Sign-on to Office, the Add-in, and Microsoft Graph

The `getAccessTokenAsync` API in Office.js enables users who are signed into Office to get access to an AAD-protected add-in and to Microsoft Graph without needing to sign-in again. This sample is built on Node.js and express. 

 > Note: The `getAccessTokenAsync` API is in preview.

## Table of Contents
* [Change History](#change-history)
* [Prerequisites](#prerequisites)
* [To use the project](#to-use-the-project)
* [Questions and comments](#questions-and-comments)
* [Additional resources](#additional-resources)

## Change History

* May 10, 2017: Initial version.
* September 15, 2017: Added handling for 2FA.
* December 8, 2017: Added extensive error handling.
* December 19, 2018: Updated to more recent versions of some dependencies.
* January 7, 2019: Added information about application security mitigations.

## Prerequisites

* An Office 365 account.
* During the preview phase, the SSO requires Office 365 (the subscription version of Office, also called “Click to Run”). You should use the latest monthly version and build from the Insiders channel. You need to be an Office Insider to get this version. For more information, see [Be an Office Insider](https://products.office.com/office-insider?tab=tab-1). Please note that when a build graduates to the production semi-annual channel, support for preview features, including SSO, is turned off for that build.
* [Git Bash](https://git-scm.com/downloads) (Or another git client.)
* TypeScript version 2.2.2 or later.

## Deviations from Best Practices

The samples in this repo are narrowly focused on demonstrating the use of the SSO APIs. To keep it simple, some best practices are not followed, including best practices in web application security. *You should not use any of these samples as the starting point of a production add-in unless you are prepared to make substantial changes.* We recommend that you begin a production add-in by using one of the Office Add-in projects in Visual Studio, or by generating a new project with the [Yeoman Generator for Office Add-ins](https://github.com/OfficeDev/generator-office).

_Some_ of the points to keep in mind about these samples:

* Do not ship reusable certs as these samples do. Produce your own certs for your server and make sure they are not web-accessible.
* The samples send a hardcoded query parameter on the URL for the Microsoft Graph REST API. If you modify this code in a production add-in and any part of query parameter comes from user input, be sure that it is sanitized so that it cannot be used in a Response header injection attack.

## To use the project

This sample is meant to accompany the walkthrough at: [Create a Node.js Office Add-in that uses Single Sign-on (preview)](https://dev.office.com/docs/add-ins/develop/create-sso-office-add-ins-nodejs).

There are three versions of the sample, in the folders **Before**, **Completed**, **Completed Multitenant**.

To use the Before version and manually add the crucial SSO-oriented code, follow all the procedures in the article linked to above.

To work with the Completed versions, follow all the procedures, except the sections "Code the client-side" and "Code the server-side" in the article linked to above.

_Completed Multitenant_ version allows you to use SSO with any Microsoft account regardless of its domain.

> **IMPORTANT**: Regardless of which version you use, you will need to trust a certificate for the localhost. Follow the instructions [here](https://github.com/OfficeDev/generator-office/blob/master/src/docs/ssl.md), except that the `certs` folders for each of the versions in this repo is in the `/src` folder, not the root folder.

## Questions and comments

We'd love to get your feedback about this sample. You can send your feedback to us in the *Issues* section of this repository.

Questions about Microsoft Office 365 development in general should be posted to [Stack Overflow](http://stackoverflow.com/questions/tagged/office-js+API). If your question is about the Office JavaScript APIs, make sure that your questions are tagged with [office-js] and [API].

## Additional resources

* [Office add-in documentation](https://msdn.microsoft.com/en-us/library/office/jj220060.aspx)
* [Office Dev Center](http://dev.office.com/)
* More Office Add-in samples at [OfficeDev on Github](https://github.com/officedev)

## Copyright

Copyright (c) 2017 Microsoft Corporation. All rights reserved.

This project has adopted the [Microsoft Open Source Code of Conduct](https://opensource.microsoft.com/codeofconduct/). For more information, see the [Code of Conduct FAQ](https://opensource.microsoft.com/codeofconduct/faq/) or contact [opencode@microsoft.com](mailto:opencode@microsoft.com) with any additional questions or comments.
