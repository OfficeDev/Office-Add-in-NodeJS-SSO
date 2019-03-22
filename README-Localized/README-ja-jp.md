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
# <a name="office-add-in-that-supports-single-sign-on-to-office-the-add-in-and-microsoft-graph"></a>Office、アドイン、Microsoft Graph へのシングル サインオンをサポートする Office アドイン

Office.js で `getAccessTokenAsync` API を使用すると、Office にサインインしたユーザーは、再サインインすることなく、AAD 保護アドインと Microsoft Graph にアクセスできます。 このサンプルは Node.js と Express に基づいてビルドされています。 

 > 注意:`getAccessTokenAsync` API はプレビュー段階です。

## <a name="table-of-contents"></a>目次
* [変更履歴](#change-history)
* [前提条件](#prerequisites)
* [プロジェクトを使用するには](#to-use-the-project)
* [質問とコメント](#questions-and-comments)
* [その他のリソース](#additional-resources)

## <a name="change-history"></a>変更履歴

* 2017 年 5 月 10 日:初期バージョン。
* 2017 年 9 月 15 日:2FA の処理が追加されました。
* 2017 年 12 月 8 日:広範なエラー処理が追加されました。
* 2018 年 12 月 19 日:一部の依存関係が最新バージョンに更新されました。
* 2019 年 1 月 7 日:アプリケーションのセキュリティ対策に関する情報が追加されました。

## <a name="prerequisites"></a>前提条件

* Office 365 アカウント。
* プレビュー フェーズ中、SSO には、Office 365 (“クイック実行“ と呼ばれることもある Office のサブスクリプション バージョン) が必要です。 Insider チャネルからの最新の月次バージョンとビルドを使ってください。 このバージョンを入手するには、Office Insider への参加が必要です。 詳細については、「[Office Insider になる](https://products.office.com/office-insider?tab=tab-1)」を参照してください。 ビルドが半期チャネルの運用に移行すると、そのビルドで SSO を含むプレビュー機能のサポートはオフになりますので、ご注意ください。
* [Git Bash](https://git-scm.com/downloads) (またはその他の Git クライアント。)
* TypeScript バージョン 2.2.2 以降。

## <a name="deviations-from-best-practices"></a>ベスト プラクティスからの逸脱

このリポジトリのサンプルでは、SSO API の使用方法を示すことに焦点を絞っています。 わかりやすくするため、Web アプリケーション セキュリティのベスト プラクティスを含む、いくつかのベスト プラクティスに従っていません。 *大幅に変更する準備ができていない場合は、いずれのサンプルも、運用環境のアドインのベースとして使用しないでください。* Visual Studio のいずれかの Office アドイン プロジェクトをベースにして運用環境のアドインを開始するか、[Office アドインの Yeoman ジェネレーター](https://github.com/OfficeDev/generator-office)で新しいプロジェクトを作成することをお勧めします。

これらのサンプルに関する_いくつか_の注意事項:

* これらのサンプルのように再利用可能な証明書を送付しないでください。 サーバーに対して独自の証明書を作成し、Web でアクセス可能ではないことを確認します。
* サンプルでは、Microsoft Graph REST API の URL のハードコーディングされたクエリ パラメーターを送信します。 運用環境のアドインでこのコードを変更し、クエリ パラメーターの一部分がユーザーの入力に基づいている場合、このコードがサニタイズされ、応答ヘッダーの挿入攻撃で使用できないようになっていることを確認してください。

## <a name="to-use-the-project"></a>プロジェクトを使用するには

このサンプルは、次のチュートリアルに添付するためのものです。[シングル サインオンを使用する Node.js Office アドインを作成する (プレビュー)](https://dev.office.com/docs/add-ins/develop/create-sso-office-add-ins-nodejs)。

フォルダーには、**Before**、**Completed**、**Completed Multitenant** の 3 つのバージョンのサンプルがあります。

Before バージョンを使用して重要な SSO 指向コードを手動で追加するには、上にリンクされている記事のすべての手順に従います。

Completed バージョンを操作するには、上にリンクされている記事の "クライアント側のコードを作成する" セクションと "サーバー側のコードを作成する" セクション以外の手順に従います。

_Completed Multitenant_ バージョンでは、そのドメインに関係なく、すべての Microsoft アカウントで SSO を使用することができます。

> **重要**:ローカル ホストの証明書は、使用するバージョンにかかわらず信頼する必要があります。 [ここ](https://github.com/OfficeDev/generator-office/blob/master/src/docs/ssl.md)の手順に従います。ただし、このリポジトリの各バージョンの `certs` フォルダーは、ルート フォルダーではなく `/src` フォルダー内にあります。

## <a name="questions-and-comments"></a>質問とコメント

このサンプルに関するフィードバックをお寄せください。このリポジトリの「*問題*」セクションでフィードバックを送信できます。

Microsoft Office 365 開発全般の質問につきましては、「[スタック オーバーフロー](http://stackoverflow.com/questions/tagged/office-js+API)」に投稿してください。Office JavaScript API に関する質問の場合は、必ず質問に [office-js] と [API] のタグを付けてください。

## <a name="additional-resources"></a>追加リソース

* [Office アドインのドキュメント](https://msdn.microsoft.com/ja-jp/library/office/jj220060.aspx)
* [Office デベロッパー センター](http://dev.office.com/)
* [Github の OfficeDev](https://github.com/officedev) にあるその他の Office アドイン サンプル

## <a name="copyright"></a>著作権

Copyright (c) 2017 Microsoft Corporation.All rights reserved.

このプロジェクトでは、[Microsoft Open Source Code of Conduct](https://opensource.microsoft.com/codeofconduct/) が採用されています。詳細については、「[Code of Conduct の FAQ](https://opensource.microsoft.com/codeofconduct/faq/)」を参照してください。また、その他の質問やコメントがあれば、[opencode@microsoft.com](mailto:opencode@microsoft.com) までお問い合わせください。
