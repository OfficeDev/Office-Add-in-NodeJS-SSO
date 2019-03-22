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
# <a name="office-add-in-that-supports-single-sign-on-to-office-the-add-in-and-microsoft-graph"></a>Complément d’Office qui prend en charge authentification unique pour Office, le complément et Microsoft Graph

L’API`getAccessTokenAsync` dans Office.js permet aux utilisateurs qui sont connectés à Office d’accéder à un complément protégé par AAD et à Microsoft Graph sans avoir à se reconnecter. Cet exemple repose sur Node.js et Express. 

 > Remarque : Cette API`getAccessTokenAsync` est disponible en aperçu.

## <a name="table-of-contents"></a>Sommaire
* [Historique des modifications](#change-history)
* [Conditions préalables](#prerequisites)
* [Utiliser l’explorateur de projets](#to-use-the-project)
* [Questions et commentaires](#questions-and-comments)
* [Ressources supplémentaires](#additional-resources)

## <a name="change-history"></a>Historique des modifications

* 10 mai 2017 : Version d’origine.
* 15 Septembre 2017: Gestion ajoutée pour 2FA.
* 8 décembre 2017: Gestion ajouté des erreurs étendus.
* 19 décembre 2018: Mise à jour vers les versions plus récentes de certaines dépendances.
* 7 janvier 2019: Ajout d’informations sur les atténuations de sécurité d’application.

## <a name="prerequisites"></a>Conditions requises

* Un compte Office 365.
* L’authentification unique SSO requiert Office 365 (la version par abonnement d’Office, également appelée « Démarrer en un clic »). Vous devez utiliser la version et le build mensuels les plus récents du canal du programme Insider. Vous devez participer au programme Office Insider pour obtenir cette version. Pour plus d’informations, reportez-vous à [Participez au programme Office Insider](https://products.office.com/office-insider?tab=tab-1). Veuillez noter que lorsqu’un build passe au canal semi-annuel de production, la prise en charge des fonctionnalités d’aperçu, y compris l’authentification unique, est désactivée pour ce build.
* [GIT Bash](https://git-scm.com/downloads) (ou un autre client Git)
* TypeScript version 2.2.2 ou ultérieure.

## <a name="deviations-from-best-practices"></a>Écarts entre les meilleures pratiques

Les exemples dans cette repo sont étroitement axées sur la démonstration de l’utilisation de l’API d’authentification unique SSO. Pour effectuer une opération simple, certaines pratiques recommandées ne sont pas suivies, y compris les meilleures pratiques de sécurité de l’application web. *Vous ne devez pas utiliser un de ces exemples comme point de départ du complément production, sauf si vous êtes prêt à apporter des modifications substantielles.* Nous vous recommandons de commencer un complément production en utilisant l’un des projets complément Office dans Visual Studio ou en générer un nouveau projet avec la [générateur Yeoman pour les compléments Office](https://github.com/OfficeDev/generator-office).

_Certains_ des points à retenir concernant ces exemples :

* Ne livrent pas de certificats réutilisables comme ces exemples. Génèrent vos propres certificats par pour votre serveur et assurez-vous qu’ils ne sont pas accessibles sur le web.
* Les exemples envoient un paramètre de requête codée en dur dans l’URL pour le Microsoft Graph l’API REST. Si vous modifiez ce code dans un complément production et une partie quelconque de paramètre de requête provient d’une intervention de l’utilisateur, n’oubliez pas qu’il est purgé afin qu’il ne puisse pas être utilisé dans une attaque par injection d’en-tête de réponse.

## <a name="to-use-the-project"></a>Utiliser le projet

Cet exemple est destiné à accompagnent cette procédure en : [Créer un complément Office Node.js qui utilise l’authentification unique (aperçu)](https://dev.office.com/docs/add-ins/develop/create-sso-office-add-ins-nodejs).

Il existe trois versions de l’échantillon dans les dossiers **Avant**, **Terminée**, **Terminée pouvant être partagée**.

Pour utiliser la version précédente et ajouter manuellement le code de l’authentification unique orientée essentiel, suivre les procédures décrites dans l’article lié à ci-dessus.

Pour travailler avec les versions terminée, suivez les procédures, sauf les sections « Code côté client » et « Code côté serveur » dans l’article lié à ci-dessus.

_Terminé pouvant être partagée_ version vous autorise à utiliser l’authentification unique SSO avec un compte Microsoft, quel que soit son domaine.

> **IMPORTANT**: Quelle que soit la version que vous utilisez, vous devrez approuver un certificat pour l’hôte local. Suivez les instructions [ici](https://github.com/OfficeDev/generator-office/blob/master/src/docs/ssl.md), sauf que les dossiers pour chacune des versions dans cette repo est dans le  dossier, pas dans le dossier racine.

## <a name="questions-and-comments"></a>Questions et commentaires

Nous serions ravis de connaître votre opinion sur cet exemple. Vous pouvez nous envoyer vos commentaires via la section *Problèmes* de ce référentiel.

Les questions générales sur le développement de Microsoft Office 365 doivent être publiées sur [Stack Overflow](http://stackoverflow.com/questions/tagged/office-js+API). Si votre question concerne les API Office JavaScript, assurez-vous qu’elle comporte les balises [office-js] et [API].

## <a name="additional-resources"></a>Ressources supplémentaires

* [Documentation de complément Office](https://msdn.microsoft.com/fr-fr/library/office/jj220060.aspx)
* [Centre de développement Office](http://dev.office.com/)
* Plus d’exemples de complément Office sur [OfficeDev sur Github](https://github.com/officedev)

## <a name="copyright"></a>Copyright

Copyright (c) 2017 Microsoft Corporation. Tous droits réservés.

Ce projet a adopté le [code de conduite Open Source de Microsoft](https://opensource.microsoft.com/codeofconduct/). Pour plus d’informations, reportez-vous à la [FAQ relative au code de conduite](https://opensource.microsoft.com/codeofconduct/faq/) ou contactez [opencode@microsoft.com](mailto:opencode@microsoft.com) pour toute question ou tout commentaire.
