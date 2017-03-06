# <a name="office-add-in-that-converts-directly-between-word-and-markdown-formats"></a>直接转换 Word 和 Markdown 格式的 Office 外接程序

使用 Word.js API 将 Markdown 文档转换成 Word 格式以供编辑，然后使用 Paragraph、Table、List 和 Range 对象将 Word 文档转换回 Markdown 格式。

![转换 Word 和 Markdown 格式](readme_art/ReadMeScreenshot.PNG)

## <a name="table-of-contents"></a>目录
* [修订记录](#change-history)
* [先决条件](#prerequisites)
* [测试外接程序](#test-the-add-in)
* [已知问题](#known-issues)
* [问题和意见](#questions-and-comments)
* [其他资源](#additional-resources)

## <a name="change-history"></a>修订记录

2016 年 12 月 16 日：

* 首版。

## <a name="prerequisites"></a>先决条件

* Visual Studio 2015 或更高版本。
* Word 2016 for Windows（内部版本 16.0.6727.1000 或更高版本）。

## <a name="test-the-add-in"></a>测试外接程序

1. 将项目克隆或下载到桌面。
2. 在 Visual Studio 中，打开 Word-Add-in-JavaScript-MDConversion.sln 文件。
2. 按 F5。
3. 在 Word 启动后，按“**主页**”功能区上的“**打开转换器**”按钮。
4. 在应用程序已加载后，按“**插入测试 Markdown 文档**”按钮。
5. 在示例 Markdown 文本已加载后，按“**将 MD 文本转换成 Word**”按钮。
6. 在将文档转换成 Word 格式后，编辑文档。 
7. 按“**将文档转换成 Markdown 格式**”按钮。 
8. 转换文档后，将文档内容复制并粘贴到 Markdown 预览器（如 Visual Studio Code）中。
9. 也可以先按“**插入测试 Word 文档**”按钮，然后将创建的示例 Word 文档转换成 Markdown 格式。 
10. （可选）从你自己的 Markdown 文本或 Word 内容入手，测试外接程序。

## <a name="known-issues"></a>已知问题

- 由于以编程方式创建的 Word 列表的创建方式有 bug，因此在转换 Markdown 和 Word 格式时，只能正确转换文档中的第一个列表（或有时为前两个列表）。（所有 Markdown 列表都会正确转换成 Word 格式。）
- 如果对同一文档重复来回转换 Word 和 Markdown 格式，表中的所有行都会采用标题行格式，这通常包括粗体文本。
- 外接程序会使用 Word Online 尚不支持的部分 Office API（截至 2017 年 2 月 15 日）。应在桌面 Word 中对其进行测试（在按 F5 时它将自动打开。

## <a name="questions-and-comments"></a>问题和意见

我们乐意倾听你对此示例的反馈。你可以在此存储库中的“*问题*”部分向我们发送反馈。

与 Microsoft Office 365 开发相关的一般问题应发布到 [Stack Overflow](http://stackoverflow.com/questions/tagged/office-js+API)。如果你的问题是关于 Office JavaScript API，请务必为问题添加 [office-js] 和 [API].标记。

## <a name="additional-resources"></a>其他资源

* [Office 外接程序文档](https://msdn.microsoft.com/en-us/library/office/jj220060.aspx)
* [Office 开发人员中心](http://dev.office.com/)
* 有关更多 Office 外接程序示例，请访问 [Github 上的 OfficeDev](https://github.com/officedev)。

## <a name="copyright"></a>版权
版权所有 © 2016 Microsoft Corporation。保留所有权利。

