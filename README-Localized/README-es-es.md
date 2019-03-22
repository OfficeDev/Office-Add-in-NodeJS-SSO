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
# <a name="office-add-in-that-supports-single-sign-on-to-office-the-add-in-and-microsoft-graph"></a>Complemento de Office compatible con el inicio de sesión único de Office, el complemento y Microsoft Graph

La API de `getAccessTokenAsync` en Office.js permite a los usuarios que hayan iniciado sesión en Office obtener acceso a un complemento protegido por AAD y a Microsoft Graph sin necesidad de volver a iniciar sesión. Este ejemplo se basa en Node.js y Express. 

 > Nota: La API de `getAccessTokenAsync` está en versión preliminar.

## <a name="table-of-contents"></a>Tabla de contenido
* [Historial de cambios](#change-history)
* [Requisitos previos](#prerequisites)
* [Usar el proyecto](#to-use-the-project)
* [Preguntas y comentarios](#questions-and-comments)
* [Recursos adicionales](#additional-resources)

## <a name="change-history"></a>Historial de cambios

* 10 de mayo de 2017: Versión inicial.
* 15 de septiembre de 2017: Agregado el control de 2FA.
* 8 de diciembre de 2017: Agregado amplio control de errores.
* 19 de diciembre de 2018: Actualizado a versiones más recientes de algunas dependencias.
* 7 de enero de 2019: Agregada información sobre mitigaciones de la seguridad de aplicaciones.

## <a name="prerequisites"></a>Requisitos previos

* Una cuenta de Office 365.
* Durante la fase de versión preliminar, el SSO requiere Office 365 (la versión de suscripción de Office, también denominada "Hacer clic y ejecutar"). Debería usar la última versión y compilación mensual del canal Insider. Necesita ser participante de Office Insider para obtener esta versión. Para más información, vea [Participar en Office Insider](https://products.office.com/office-insider?tab=tab-1). Tenga en cuenta que cuando una compilación pasa al canal de producción semianual, la compatibilidad para las características de vista previa, incluido el inicio de sesión único, se desactivan para esa versión.
* [Git Bash](https://git-scm.com/downloads) (U otro cliente de Git).
* TypeScript versión 2.2.2 o posterior.

## <a name="deviations-from-best-practices"></a>Desviaciones de los procedimientos recomendados

Los ejemplos de este repositorio se centran concretamente en mostrar el uso de la API de SSO. Para hacerlo fácil, algunos procedimientos recomendados no se siguen, incluidas mejores prácticas de seguridad de la aplicación web. *No debería usar estos ejemplos como punto de partida de un complemento de producción a menos que esté preparado para realizar cambios importantes.* Le recomendamos que empiece con un complemento de producción con uno de los proyectos de complemento de Office en Visual Studio o creando un nuevo proyecto con el [Generador Yeoman para complementos de Office](https://github.com/OfficeDev/generator-office).

_Algunos_ de los aspectos a tener en cuenta acerca de estos ejemplos:

* No envíe certificados reutilizables como se hace en estos ejemplos. Cree sus propios certificados para el servidor y asegúrese de que no son accesibles para la web.
* Los ejemplos envían un parámetro de la consulta codificado de forma rígida en la dirección URL de la API de REST de Microsoft Graph. Si modifica este código en un complemento de producción y algunas partes de los parámetros de consulta proceden de las entradas del usuario, asegúrese de que se depuran para que no pueda usarse en un ataque de inyección de encabezado de respuesta.

## <a name="to-use-the-project"></a>Usar el proyecto

Este ejemplo se usa para acompañar el tutorial en: [Crear un complemento de Node.js Office que usa el inicio de sesión único (versión preliminar)](https://dev.office.com/docs/add-ins/develop/create-sso-office-add-ins-nodejs)

Existen dos versiones del ejemplo en las carpetas **Anterior**, **Completada** y **Completada multiinquilino**.

Para usar la versión Anterior y agregar manualmente el código esencial orientado SSO, siga los procedimientos descritos en el artículo vinculado anteriormente.

Para trabajar con las versiones Completas, siga todos los procedimientos, excepto las secciones "Código del cliente" y "Código del lado servidor" en el artículo vinculado anteriormente.

_Completada Multiinquilino_: esta versión le permite usar SSO con cualquier cuenta de Microsoft, independientemente de su dominio.

> **IMPORTANTE**: Independientemente de la versión que use, debe confiar en un certificado para el host local. Siga [estas instrucciones](https://github.com/OfficeDev/generator-office/blob/master/src/docs/ssl.md), pero tenga en cuenta que las carpetas de `certs` para cada una de las versiones en este repo están en la carpeta `/src`, no la carpeta raíz.

## <a name="questions-and-comments"></a>Preguntas y comentarios

Nos encantaría recibir sus comentarios sobre este ejemplo. Puede enviarnos comentarios a través de la sección *Problemas* de este repositorio.

Las preguntas generales sobre el desarrollo de Microsoft Office 365 deben publicarse en [Stack Overflow](http://stackoverflow.com/questions/tagged/office-js+API). Si su pregunta trata sobre las API de JavaScript para Office, asegúrese de que se etiqueta con [office-js] y [API].

## <a name="additional-resources"></a>Recursos adicionales

* [Documentación de complementos de Office](https://msdn.microsoft.com/es-es/library/office/jj220060.aspx)
* [Centro de desarrollo de Office](http://dev.office.com/)
* Más ejemplos de complementos de Office en [OfficeDev en GitHub](https://github.com/officedev)

## <a name="copyright"></a>Derechos de autor

Copyright (c) 2017 Microsoft Corporation. Todos los derechos reservados.

Este proyecto ha adoptado el [Código de conducta de código abierto de Microsoft](https://opensource.microsoft.com/codeofconduct/). Para obtener más información, consulte las [preguntas más frecuentes sobre el Código de conducta](https://opensource.microsoft.com/codeofconduct/faq/) o póngase en contacto con [opencode@microsoft.com](mailto:opencode@microsoft.com) si tiene otras preguntas o comentarios.
