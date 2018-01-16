# <a name="office-add-in-that-converts-directly-between-word-and-markdown-formats"></a>Office-Add-in, das eine direkt Konvertierung zwischen Word- und Markdown-Formaten durchführt

Verwenden Sie die Word.js-APIs, um ein Markdown-Dokument in Word zu konvertieren, und konvertieren Sie dann das Word-Dokument mithilfe der Paragraph-, Table, List- und Range-Objekte zurück in das Markdown-Format.

![Konvertierung zwischen Word und Markdown](readme_art/ReadMeScreenshot.PNG)

## <a name="table-of-contents"></a>Inhalt
* [Änderungsverlauf](#change-history)
* [Voraussetzungen](#prerequisites)
* [Testen des Add-Ins](#test-the-add-in)
* [Bekannte Probleme](#known-issues)
* [Fragen und Kommentare](#questions-and-comments)
* [Zusätzliche Ressourcen](#additional-resources)

## <a name="change-history"></a>Änderungsverlauf

16. Dezember 2016

* Ursprüngliche Version

## <a name="prerequisites"></a>Voraussetzungen

* Visual Studio 2015 oder höher.
* Word 2016 für Windows, Build 16.0.6727.1000 oder höher

## <a name="test-the-add-in"></a>Testen des Add-Ins

1. Klonen Sie das Projekt auf Ihrem Desktop, oder laden Sie es herunter.
2. Öffnen Sie die Datei „Word-Add-in-JavaScript-MDConversion.sln“ in Visual Studio.
2. Drücken Sie F5.
3. Nachdem Word gestartet wurde, drücken Sie die Schaltfläche **Konverter öffnen** Schaltfläche auf dem Menüband **Start**.
4. Wenn die Anwendung geladen wurde, drücken Sie die Schaltfläche **Markdown-Testdokument einfügen**.
5. Nachdem der Beispiel-Markdowntext geladen wurde, drücken Sie die Schaltfläche **MD-Text in Word konvertieren**.
6. Bearbeiten Sie das Dokument, nachdem es in Word konvertiert wurde. 
7. Drücken Sie die Schaltfläche **Dokument in Markdown konvertieren**. 
8. Nachdem das Dokument konvertiert wurde, kopieren Sie seinen Inhalt in eine Markdown-Vorschau, z. B. Visual Studio Code.
9. Alternativ können Sie mit der Schaltfläche **Word-Testdokument einfügen** beginnen und das erstellte Word-Beispieldokument in Markdown konvertieren. 
10. Beginnen Sie optional mit Ihrem eigenen Markdown-Text oder Word-Inhalt, und testen Sie das Add-In.

## <a name="known-issues"></a>Bekannte Probleme

- Aufgrund eines Fehlers in der Art und Weise, wie programmgesteuert erstellte Word-Listen erstellt werden, wird mit Markdown-to-Word nur die erste Liste (oder manchmal die beiden ersten Listen) in einem Dokument korrekt konvertiert. (Eine beliebige Anzahl von Markdown-Listen wird korrekt in Word konvertiert.)
- Wenn Sie dasselbe Dokument wiederholt zwischen Word und Markdown hin und her konvertieren, verwenden alle Zeilen in der Tabelle die Formatierung der Kopfzeile, die in der Regel fett formatierten Text enthält.
- Das Add-In verwendet einige Office-APIs, die in Word Online noch nicht unterstützt werden (Stand 15.02.2017). Sie sollten es in der Word-Desktopanwendung testen (die automatisch geöffnet wird, wenn Sie F5 drücken).

## <a name="questions-and-comments"></a>Fragen und Kommentare

Wir schätzen Ihr Feedback hinsichtlich dieses Beispiels. Sie können uns Ihr Feedback über den Abschnitt *Probleme* dieses Repositorys senden.

Fragen zur Microsoft Office 365-Entwicklung sollten in [Stack Overflow](http://stackoverflow.com/questions/tagged/office-js+API) gestellt werden. Wenn Ihre Frage die Office JavaScript-APIs betrifft, sollte die Frage mit [office-js] und [API] kategorisiert sein.

## <a name="additional-resources"></a>Zusätzliche Ressourcen

* 
  [Dokumentation zu Office-Add-Ins](https://msdn.microsoft.com/en-us/library/office/jj220060.aspx)
* [Office Dev Center](http://dev.office.com/)
* Weitere Office-Add-In-Beispiele unter [OfficeDev auf Github](https://github.com/officedev)

## <a name="copyright"></a>Copyright
Copyright (c) 2016 Microsoft Corporation. Alle Rechte vorbehalten.



In diesem Projekt wurden die [Microsoft Open Source-Verhaltensregeln](https://opensource.microsoft.com/codeofconduct/) übernommen. Weitere Informationen finden Sie unter [Häufig gestellte Fragen zu Verhaltensregeln](https://opensource.microsoft.com/codeofconduct/faq/), oder richten Sie Ihre Fragen oder Kommentare an [opencode@microsoft.com](mailto:opencode@microsoft.com).
