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
# <a name="office-add-in-that-supports-single-sign-on-to-office-the-add-in-and-microsoft-graph"></a>Suplemento Office compatível com logon único para o Office, o suplemento e Microsoft Graph

A `getAccessTokenAsync` API no Office.js permite aos usuários que estão conectados ao Office para obter acesso a um suplemento protegido pelo AAD e ao Microsoft Graph sem precisar entrar novamente. Este exemplo é criado em Node.js e express. 

 > Observação: A `getAccessTokenAsync` API está em visualização.

## <a name="table-of-contents"></a>Sumário
* [Histórico de Alterações](#change-history)
* [Pré-requisitos](#prerequisites)
* [Usar o projeto](#to-use-the-project)
* [Perguntas e comentários](#questions-and-comments)
* [Recursos adicionais](#additional-resources)

## <a name="change-history"></a>Histórico de alterações

* 10 de maio de 2017: Versão inicial.
* 15 de setembro de 2017: Adicionar o tratamento de 2FA.
* 8 de dezembro de 2017: Adicionar o tratamento de erro amplo.
* 19 de dezembro de 2018: Atualizado para versões mais recentes de algumas dependências.
* 7 de janeiro de 2019: Informações adicionais sobre atenuações de segurança do aplicativo.

## <a name="prerequisites"></a>Pré-requisitos

* Uma conta do Office 365.
* Durante a fase de visualização, O SSO requer o Office 365 (versão de assinatura do Office, também chamada "Clique para Executar"). Você deve usar o build e a versão mensal mais recente do canal Insiders. É necessário ingressar no programa Office Insider para obter essa versão. Para saber mais, confira a página [Seja um Office Insider](https://products.office.com/office-insider?tab=tab-1). Observe que, quando um build é promovido ao Canal Semestral de produção, o suporte para recursos de visualização, como o SSO, é desativado para esse build.
* [Git Bash](https://git-scm.com/downloads) (ou outro cliente Git.)
* TypeScript, versão 2.2.2 ou posterior.

## <a name="deviations-from-best-practices"></a>Desvios de práticas recomendadas

Os exemplos neste repositório são estritamente concentrados em demonstrar o uso de APIs SSO. Para manter a simplicidade, algumas práticas recomendadas não foram seguidas, incluindo as práticas recomendadas em segurança do aplicativo web. *Você não deve usar qualquer um desses exemplos como ponto de partida de um suplemento de produção, a menos que você esteja preparado para fazer alterações significativas.* Recomendamos começar um suplemento de produção, use um dos projetos suplementos do Office no Visual Studio ou gerar um novo projeto com o [gerador Yeoman para suplementos do Office](https://github.com/OfficeDev/generator-office).

_Alguns_ dos pontos importantes sobre esses exemplos:

* Não é distribuído certificados reutilizáveis como esses exemplos. Produz seus próprios certificados para seu servidor e garante que não estão acessíveis a web.
* Os exemplos enviam um parâmetro de consulta codificada na URL para a API REST do Microsoft Graph. Se você modificar o código em um suplemento de produção e provenientes de qualquer parte do parâmetro de consulta de entrada do usuário, certifique-se de que estão limpos para que não possam ser usados em um ataque de inserção de cabeçalho de resposta.

## <a name="to-use-the-project"></a>Usar o projeto

Este exemplo deve acompanhar o passo a passo em: [Crie um Suplemento do Office com Node.js que use logon único (prévia)](https://dev.office.com/docs/add-ins/develop/create-sso-office-add-ins-nodejs).

Há três versões da amostra nas pastas **antes**, **concluída**, **multilocatário concluído**.

Para usar a versão de antes e adicionar manualmente o código fundamental orientado SSO, siga todos os procedimentos do artigo acima.

Para trabalhar com as versões concluídas, siga os procedimentos, exceto as seções "Código do cliente" e "Do servidor do código" no artigo vinculado a acima.

_Multilocatário Concluído_ versão permite que você use o SSO com uma conta da Microsoft independentemente do seu domínio.

> **IMPORTANTE**: Independentemente de qual versão você usa, será necessário confiar em um certificado para um host local. Siga as instruções [aqui](https://github.com/OfficeDev/generator-office/blob/master/src/docs/ssl.md), exceto que as `certs` pastas para cada uma das versões deste repositório está na `/src` pasta, não na pasta raiz.

## <a name="questions-and-comments"></a>Perguntas e comentários

Gostaríamos de saber sua opinião sobre este exemplo. Você pode nos enviar comentários na seção *Problemas* deste repositório.

As perguntas sobre o desenvolvimento do Microsoft Office 365 em geral devem ser postadas no [Stack Overflow](http://stackoverflow.com/questions/tagged/office-js+API). Se sua pergunta estiver relacionada às APIs JavaScript para Office, não deixe de marcá-la com as tags [office-js] e [API].

## <a name="additional-resources"></a>Recursos adicionais

* [Documentação dos suplementos do Office](https://msdn.microsoft.com/pt-br/library/office/jj220060.aspx)
* [Centro de Desenvolvimento do Office](http://dev.office.com/)
* Confira outros exemplos de Suplemento do Office em [OfficeDev no Github](https://github.com/officedev)

## <a name="copyright"></a>Direitos autorais

Copyright (c) 2017 Microsoft Corporation. Todos os direitos reservados.

Este projeto adotou o [Código de Conduta de Software Livre da Microsoft](https://opensource.microsoft.com/codeofconduct/). Saiba mais nas [Perguntas frequentes sobre o Código de Conduta](https://opensource.microsoft.com/codeofconduct/faq/) ou contate [opencode@microsoft.com](mailto:opencode@microsoft.com) se tiver outras dúvidas ou comentários.
