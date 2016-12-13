
# A Sample Document (first level header)

This document has it all; **bolded** text, *italicized* text, bookmark links, a web link, a table, a code block, and a very complicated list.

This is another paragraph, because you can't have *too* many.

## Table of Contents with Bookmark Links (second level header)

[Table](#table-second-level-header) 

[Code Block](#code-block-fourth-level-header) 

[Paragraph with Web Link](#paragraph-with-web-link-third-level-header) 

[Complex List](#complex-list-fifth-level-header) 

## Table (second level header)

| Name | Mood | Hobby | 
|:--|:--|:--|
| Bob | happy | Cooking and baking | 
| Sally | wistful | Running marathons **barefoot** | 
| Mary | energetic | Playing multiple games of chess simultaneously | 

#### Code Block (fourth level header)

```
// This code is only here to illustrate a code block.
Word.run(function (context) {
    var body = context.document.body;
    body.clear();
    return context.sync().then(function () {
        console.log('Cleared the body contents.');
    });
})
```

### Paragraph with Web Link (third level header)

For more information about Office host application and server requirements, see [Requirements for running Office Add-ins](https://msdn.microsoft.com/EN-US/library/office/dn833104.aspx). 

##### Complex List (fifth level header)

There can be up to nine levels in a list.

* Top level is bullet
    * Second level is bullet
    * Another second level item
        1. Third level is numeric
            1. Fourth level is numeric
            1. Another fourth level item.
                * Fifth level is bullet
                    1. Sixth level is numeric
                    1. Another sixth level item
                * Back to the fifth level
                    * Back to the sixth level
                        * Seventh is bullet
                            1. Eighth is numeric
                                * Ninth is bullet
                                * Another ninth level item
                            1. Another eighth level item
                        * Another seventh level
                    * Another sixth level item
            1. Another fourth level item
        1. Another third level item
    * Another second level item
* Another top level item
