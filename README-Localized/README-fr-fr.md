# <a name="office-add-in-that-converts-directly-between-word-and-markdown-formats"></a>Complément Office réalisant une conversion directement entre les formats Word et Markdown

Utilisez les API Word.js pour convertir un document Markdown en Word afin de le modifier, puis pour convertir le document Word au format Markdown, en utilisant les objets Paragraphe, Tableau, Liste et Plage.

![Conversion entre Word et Markdown](readme_art/ReadMeScreenshot.PNG)

## <a name="table-of-contents"></a>Sommaire
* [Historique des modifications](#change-history)
* [Conditions préalables](#prerequisites)
* [Test du complément](#test-the-add-in)
* [Problèmes connus](#known-issues)
* [Questions et commentaires](#questions-and-comments)
* [Ressources supplémentaires](#additional-resources)

## <a name="change-history"></a>Historique des modifications

16 décembre 2016 :

* Version d’origine.

## <a name="prerequisites"></a>Conditions préalables

* Visual Studio 2015 ou version ultérieure.
* Word 2016 pour Windows, version 16.0.6727.1000 ou ultérieure.

## <a name="test-the-add-in"></a>Test du complément

1. Clonez ou téléchargez le projet sur votre ordinateur.
2. Ouvrez le fichier Word-Add-in-JavaScript-MDConversion.sln dans Visual Studio.
2. Appuyez sur la touche F5.
3. Une fois Word démarré, appuyez sur le bouton **Ouvrir le convertisseur** situé sur le ruban **Accueil**.
4. Si l’application a chargé, appuyez sur le bouton **Insérer un document Markdown test**.
5. Après le chargement du texte Markdown d’exemple, appuyez sur le bouton **Convertir le texte MD en Word**.
6. Une fois que le document a été converti en Word, modifiez-le. 
7. Appuyez sur le bouton **Convertir le document au format Markdown**. 
8. Après que le document a été converti, copiez et collez son contenu dans un générateur d’aperçu Markdown, tel que Visual Studio Code.
9. Par ailleurs, vous pouvez commencer avec le bouton **Insérer un document Word test** et convertir le document Word d’exemple qui a été créé au format Markdown. 
10. Vous pouvez également commencer par un de vos textes Markdown ou contenus Word, et tester le complément.

## <a name="known-issues"></a>Problèmes connus

- En raison d’un bogue lors de la création des listes Word créées par programme, la conversion du format Markdown en Word fonctionnera correctement uniquement pour la première liste (ou parfois les deux premières listes) dans un document. (De nombreuses listes Markdown seront converties correctement en Word.)
- Si vous convertissez le même document à plusieurs reprises entre Word et Markdown, dans un sens et dans l’autre, toutes les lignes des tableaux adopteront la mise en forme de la ligne d’en-tête, qui comprend généralement du texte en gras.
- Le complément utilise des API Office qui ne sont pas encore prises en charge dans Word Online (à compter du 15/02/2017). Vous devez le tester dans l’application de bureau Word (qui s’ouvre automatiquement lorsque vous appuyez sur F5).

## <a name="questions-and-comments"></a>Questions et commentaires

Nous serions ravis de connaître votre opinion sur cet exemple. Vous pouvez nous envoyer vos commentaires via la section *Problèmes* de ce référentiel.

Les questions générales sur le développement de Microsoft Office 365 doivent être publiées sur [Stack Overflow](http://stackoverflow.com/questions/tagged/office-js+API). Si votre question concerne les API Office JavaScript, assurez-vous qu’elle est marquée avec les balises [office js] et [API].

## <a name="additional-resources"></a>Ressources supplémentaires

* 
  [Documentation de complément Office](https://msdn.microsoft.com/en-us/library/office/jj220060.aspx)
* [Centre de développement Office](http://dev.office.com/)
* Plus d’exemples de complément Office sur [OfficeDev sur Github](https://github.com/officedev)

## <a name="copyright"></a>Copyright
Copyright (c) 2016 Microsoft Corporation. Tous droits réservés.



Ce projet a adopté le [code de conduite Open Source de Microsoft](https://opensource.microsoft.com/codeofconduct/). Pour plus d’informations, reportez-vous à la [FAQ relative au code de conduite](https://opensource.microsoft.com/codeofconduct/faq/) ou contactez [opencode@microsoft.com](mailto:opencode@microsoft.com) pour toute question ou tout commentaire.
