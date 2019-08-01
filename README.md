# [ARCHIVED] Office Add-in that converts directly between Word and Markdown formats

**Note:** This repo is archived and no longer actively maintained. Security vulnerabilities may exist in the project, or its dependencies. If you plan to reuse or run any code from this repo, be sure to perform appropriate security checks on the code or dependencies first. Do not use this project as the starting point of a production Office Add-in. Always start your production code by using the Office/SharePoint development workload in Visual Studio, or the [Yeoman generator for Office Add-ins](https://github.com/OfficeDev/generator-office), and follow security best practices as you develop the add-in. 
Use the Word.js APIs to convert a Markdown document to Word for editing and then convert the Word document back to Markdown format, using the Paragraph, Table, List, and Range objects.

![Convert between Word and Markdown](readme_art/ReadMeScreenshot.PNG)

## Table of Contents
* [Change history](#change-history)
* [Prerequisites](#prerequisites)
* [Test the add-in](#test-the-add-in)
* [Known issues](#known-issues)
* [Questions and comments](#questions-and-comments)
* [Additional resources](#additional-resources)

## Change history

December 16, 2016:

* Initial version.

## Prerequisites

* Visual Studio 2015 or later.
* Word 2016 for Windows, build 16.0.6727.1000 or later.

## Test the add-in

1. Clone or download the project to your desktop.
2. Open the Word-Add-in-JavaScript-MDConversion.sln file in Visual Studio.
2. Press F5.
3. After Word launches, press the **Open Converter** button on the **Home** ribbon.
4. When the application has loaded, press the button **Insert test Markdown document**.
5. After the sample Markdown text has loaded, press the button **Convert MD text to Word**.
6. After the document has been converted to Word, edit it. 
7. Press the button **Convert document to Markdown**. 
8. After the document has converted, copy and paste its contents into a Markdown previewer, such as Visual Studio Code.
9. Alternatively, you can start with the button **Insert test Word document** and convert the sample Word document that is created to Markdown. 
10. Optionally, start with your own Markdown text or Word content and test the add-in.

## Known issues

- Due to a bug in the way that programmatically-created Word lists are created, the Markdown-to-Word will only correctly convert the first list (or sometimes the first two lists) in a document. (Any number of Markdown lists will convert correctly to Word.)
- If you convert the same document repeatedly between Word and Markdown back and forth, all of the rows in tables will take on the formatting of the header row, which usually includes bold text.
- The add-in uses some Office APIs that are not yet supported in Word Online (as of 2017/2/15). You should test it in desktop Word (which opens automatically when you press F5.

## Questions and comments

We'd love to get your feedback about this sample. You can send your feedback to us in the *Issues* section of this repository.

Questions about Microsoft Office 365 development in general should be posted to [Stack Overflow](http://stackoverflow.com/questions/tagged/office-js+API). If your question is about the Office JavaScript APIs, make sure that your questions are tagged with [office-js] and [API].

## Additional resources

* [Office add-in documentation](https://msdn.microsoft.com/en-us/library/office/jj220060.aspx)
* [Office Dev Center](http://dev.office.com/)
* More Office Add-in samples at [OfficeDev on Github](https://github.com/officedev)

## Copyright
Copyright (c) 2016 Microsoft Corporation. All rights reserved.



This project has adopted the [Microsoft Open Source Code of Conduct](https://opensource.microsoft.com/codeofconduct/). For more information, see the [Code of Conduct FAQ](https://opensource.microsoft.com/codeofconduct/faq/) or contact [opencode@microsoft.com](mailto:opencode@microsoft.com) with any additional questions or comments.
