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
# <a name="office-add-in-that-supports-single-sign-on-to-office-the-add-in-and-microsoft-graph"></a>Office-Add-In, das einmaliges Anmelden bei Office, beim Add-In und bei Microsoft Graph unterstützt

Die `getAccessTokenAsync`-API in Office.js ermöglicht, dass Benutzer, die bei Office angemeldet sind, Zugriff auf ein durch AAD geschütztes Add-In und auf Microsoft Graph erhalten, ohne sich erneut anmelden zu müssen. Dieses Beispiel basiert auf Node.js und Express. 

 > Hinweis: Die `getAccessTokenAsync`-API befindet sich in der Vorschau.

## <a name="table-of-contents"></a>Inhaltsverzeichnis
* [Änderungsverlauf](#change-history)
* [Voraussetzungen](#prerequisites)
* [Verwenden des Projekts](#to-use-the-project)
* [Fragen und Kommentare](#questions-and-comments)
* [Zusätzliche Ressourcen](#additional-resources)

## <a name="change-history"></a>Änderungsverlauf

* 10. Mai 2017: Ursprüngliche Version.
* 15. September 2017: Verarbeitung für 2FA hinzugefügt.
* 8. Dezember 2017: Umfassende Fehlerbehandlung hinzugefügt.
* 19. Dezember 2018: Update auf neuere Versionen einiger Abhängigkeiten.
* 7. Januar 2019: Informationen zu Schutzmaßnahmen für die Anwendungssicherheit hinzugefügt.

## <a name="prerequisites"></a>Voraussetzungen

* Ein Office 365-Konto.
* Während der Vorschauphase erfordert SSO Office 365 (die Abonnementversion von Office, auch als "Klick-und-Los" bezeichnet). Sie sollten die neueste monatliche Version und den neuesten monatlichen Build aus dem Insider-Kanal verwenden. Sie müssen Office-Insider sein, um diese Version nutzen zu können. Weitere Informationen finden Sie unter [Office-Insider werden](https://products.office.com/office-insider?tab=tab-1). Bitte beachten Sie Folgendes: Wenn ein Build zum halbjährlichen Produktionskanal hochgestuft wird, ist der Support für Vorschaufeatures (einschließlich SSO) für diesen Build deaktiviert.
* [Git Bash](https://git-scm.com/downloads) (oder ein anderer Git-Client.)
* TypeScript, Version 2.2.2 oder höher.

## <a name="deviations-from-best-practices"></a>Abweichungen von bewährten Methoden

Die Beispiele in diesem Repository konzentrieren sich fast ausschließlich auf die Demonstration der Verwendung der SSO-APIs. Um die Beispiele einfach zu halten, werden einige bewährte Methoden nicht befolgt, einschließlich bewährter Methoden zur Sicherheit von Webanwendungen. *Sie sollten keines dieser Beispiele als Ausgangspunkt für ein Produktions-Add-In verwenden, es sei denn, Sie sind bereit, wesentliche Änderungen vorzunehmen.* Wir empfehlen, dass Sie ein Produktions-Add-In beginnen, indem Sie eines der Office-Add-In-Projekte in Visual Studio verwenden oder indem Sie ein neues Projekt mit dem [Yeoman-Generator für Office Add-Ins](https://github.com/OfficeDev/generator-office) erstellen.

_Einige_ Punkte, die im Hinblick auf diese Beispiele beachtet werden müssen:

* Liefern Sie keine wiederverwendbaren Zertifikate aus, auch wenn die Beispiele dies zeigen. Erstellen Sie eigene Zertifikate für Ihren Server, und stellen Sie sicher, dass diese nicht über das Internet zugänglich sind.
* Die Beispiele senden einen hartcodierten Abfrageparameter für die URL für die Microsoft Graph-REST-API. Wenn Sie diesen Code in einem Produktions-Add-In ändern und ein Teil des Abfrageparameters aus Benutzereingaben stammt, stellen Sie sicher, dass er bereinigt wird, sodass er nicht in einem Angriff mit Antwortheadereinschleusung verwendet werden kann.

## <a name="to-use-the-project"></a>Verwenden des Projekts

Dieses Beispiel soll die folgende exemplarische Vorgehensweise begleiten: [Erstellen eines Node.js-Office-Add-Ins, das einmaliges Anmelden verwendet (Vorschau)](https://dev.office.com/docs/add-ins/develop/create-sso-office-add-ins-nodejs)

Es gibt drei Versionen des Beispiels in den Ordnern **Before**, **Completed** und **Completed Multitenant**.

Um die Before-Version zu verwenden und den entscheidenden SSO-orientierten Code manuell hinzuzufügen, befolgen Sie alle Verfahren im oben verlinkten Artikel.

Um mit den Completed Versionen zu arbeiten, befolgen Sie alle Verfahren mit Ausnahme der Abschnitte "Codieren der Clientseite" und "Codieren der Serverseite" im oben verlinkten Artikel.

Die Version _Completed Multitenant_ ermöglicht es Ihnen, SSO mit jedem Microsoft-Konto unabhängig von seiner Domäne zu verwenden.

> **WICHTIG**: Ganz gleich, welche Version Sie verwenden, müssen Sie einem Zertifikat für den lokalen Host vertrauen. Befolgen Sie die [hier](https://github.com/OfficeDev/generator-office/blob/master/src/docs/ssl.md) beschriebenen Anweisungen mit der Ausnahme, dass sich die `certs`-Ordner für jede der Versionen in diesem Repository im Ordner `/src` und nicht im Stammordner befinden.

## <a name="questions-and-comments"></a>Fragen und Kommentare

Wir schätzen Ihr Feedback hinsichtlich dieses Beispiels. Sie können uns Ihr Feedback über den Abschnitt *Probleme* dieses Repositorys senden.

Fragen zur Microsoft Office 365-Entwicklung sollten in [Stack Overflow](http://stackoverflow.com/questions/tagged/office-js+API) gestellt werden. Wenn Ihre Frage die Office JavaScript-APIs betrifft, sollte die Frage mit [office-js] und [API] kategorisiert sein.

## <a name="additional-resources"></a>Zusätzliche Ressourcen

* [Dokumentation zu Office-Add-Ins](https://msdn.microsoft.com/de-de/library/office/jj220060.aspx)
* [Office Dev Center](http://dev.office.com/)
* Weitere Office-Add-In-Beispiele unter [OfficeDev auf Github](https://github.com/officedev)

## <a name="copyright"></a>Copyright

Copyright (c) 2017 Microsoft Corporation. Alle Rechte vorbehalten.

In diesem Projekt wurden die [Microsoft Open Source-Verhaltensregeln](https://opensource.microsoft.com/codeofconduct/) übernommen. Weitere Informationen finden Sie unter [Häufig gestellte Fragen zu Verhaltensregeln](https://opensource.microsoft.com/codeofconduct/faq/), oder richten Sie Ihre Fragen oder Kommentare an [opencode@microsoft.com](mailto:opencode@microsoft.com).
