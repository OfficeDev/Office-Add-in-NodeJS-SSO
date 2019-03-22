---
topic: sample
products:
- Excel
- Word
- PowerPoint
- Project
- Outlook
- Office 365
languages:
- JavaScript
- TypeScript
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
# <a name="office-add-in-that-supports-single-sign-on-to-office-the-add-in-and-microsoft-graph"></a>Office 外接程序支持单一登录到 Office、外接程序和 Microsoft Graph

Office.js 中的 `getAccessTokenAsync` API 使登录到 Office 的用户可以访问受 AAD 保护的外接程序和 Microsoft Graph，而无需再次登录。 此示例基于 Node.js 和 express 构建。 

 > 注意：`getAccessTokenAsync` API 处于预览状态。

## <a name="table-of-contents"></a>目录
* [修订记录](#change-history)
* [先决条件](#prerequisites)
* [使用项目](#to-use-the-project)
* [问题和意见](#questions-and-comments)
* [其他资源](#additional-resources)

## <a name="change-history"></a>修订记录

* 2017 年 5 月 10 日：初始版本。
* 2017 年 9 月 15 日：增加了 2FA 处理。
* 2017 年 12 月 8 日：增加了广泛的错误处理。
* 2018 年 12 月 19 日：更新到某些依赖项的更新版本。
* 2019 年 1 月 7日：增加了相关应用程序安全缓解的信息。

## <a name="prerequisites"></a>先决条件

* Office 365 帐户。
* 在预览阶段，SSO 需要 Office 365（Office 的订阅版本，也称为“即点即用版本”）。 你应该使用来自预览体验成员频道的最新每月版本和内部版本。 你可能需要成为 Office 预览体验成员，才能获取此版本。 有关详细信息，请参阅[成为 Office 预览体验成员](https://products.office.com/office-insider?tab=tab-1)。 请注意，当内部版本进入生产半年频道时，将关闭对该内部版本的预览功能（包括 SSO）的支持。
* [Git Bash](https://git-scm.com/downloads)（或其他 git 客户端）。
* TypeScript 版本 2.2.2 或更高版本。

## <a name="deviations-from-best-practices"></a>与最佳做法的偏差

此存储库中的样本仅关注演示 SSO API 的使用。 为简单起见，不遵循一些最佳实践，包括 Web 应用程序安全性的最佳实践。 *除非准备进行实质性更改，否则不应将这些样本中的任何一个用作生产外接程序的起点。* 我们建议你使用 Visual Studio 中的某个 Office 外接程序项目或使用 [Office 外接程序的 Yeoman 生成器](https://github.com/OfficeDev/generator-office)生成新项目来开始生产外接程序。

关于这些示例要记住的_几_点：

* 不要像这些示例那样寄送可重复使用的证书。 为你的服务器生成自己的证书，并确保它们不可通过 Web 访问。
* 示例在 Microsoft Graph REST API 的 URL 上发送硬编码查询参数。 如果你在生产外接程序中修改此代码，并且查询参数的任何部分来自用户输入，请确保它已被清理，以便它不能用于响应标头注入攻击。

## <a name="to-use-the-project"></a>使用项目

此示例旨在与演练一起使用：[创建使用单一登录的 Node.js Office 外接程序（预览）](https://dev.office.com/docs/add-ins/develop/create-sso-office-add-ins-nodejs)。

该示例有三个版本，分别位于 **Before**、**Completed**、**Completed Multitenant** 文件夹中。

若要使用 Before 版本并手动添加关键的面向 SSO 的代码，请按照上面链接的文章中的所有步骤进行操作。

要使用 Completed 版本，请按照上面链接的文章中的所有步骤进行操作，“编写客户端代码”和“编写服务器端代码”部分除外。

_Completed Multitenant_ 版本允许你通过任何 Microsoft 帐户（不考虑其域名）使用 SSO。

> **重要提示**：无论使用哪个版本，都需要信任 localhost 的证书。 按照[此处](https://github.com/OfficeDev/generator-office/blob/master/src/docs/ssl.md)的说明操作，但此存储库中每个版本的 `certs` 文件夹位于 `/src` 文件夹中，而不是根文件夹中。

## <a name="questions-and-comments"></a>问题和意见

我们乐意倾听你对此示例的反馈。你可以在此存储库中的“*问题*”部分向我们发送反馈。

与 Microsoft Office 365 开发相关的一般问题应发布到 [Stack Overflow](http://stackoverflow.com/questions/tagged/office-js+API)。如果你的问题是关于 Office JavaScript API，请务必为问题添加 [office-js] 和 [API].标记。

## <a name="additional-resources"></a>其他资源

* 
  [Office 外接程序文档](https://msdn.microsoft.com/zh-cn/library/office/jj220060.aspx)
* [Office 开发人员中心](http://dev.office.com/)
* 有关更多 Office 外接程序示例，请访问 [Github 上的 OfficeDev](https://github.com/officedev)。

## <a name="copyright"></a>版权信息

版权所有 (c) 2017 Microsoft Corporation。保留所有权利。

此项目已采用 [Microsoft 开放源代码行为准则](https://opensource.microsoft.com/codeofconduct/)。有关详细信息，请参阅[行为准则 FAQ](https://opensource.microsoft.com/codeofconduct/faq/)。如有其他任何问题或意见，也可联系 [opencode@microsoft.com](mailto:opencode@microsoft.com)。
