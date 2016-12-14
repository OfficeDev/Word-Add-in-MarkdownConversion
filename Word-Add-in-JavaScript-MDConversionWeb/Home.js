/* 
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/// <reference path="/Scripts/FabricUI/message.banner.js" />
/// <reference path="https://appsforoffice.microsoft.com/lib/beta/hosted/office.js" />


(function () {
    "use strict";

    var WordMarkdownConversion = window.WordMarkdownConversion || {};

    // The initialize function must be defined each time a new page is loaded.
    Office.initialize = function (reason) {
        $(document).ready(function () {
            // Initialize the FabricUI notification mechanism and hide it
            var element = document.querySelector('.ms-MessageBanner');
            FabricComponents.messageBanner = new FabricComponents.MessageBanner(element);
            FabricComponents.messageBanner.hideBanner();

            if (!Office.context.requirements.isSetSupported('WordApi', '1.2')) {
                $('#subtitle').text("Oops!");
                $("#template-description").text("Sorry, this sample requires Word 2016 or later. The buttons will not convert your document.");
                return;
            }

            $('#convert-to-markdown').click(WordMarkdownConversion.convertToMarkdown);
            $('#convert-to-word').click(WordMarkdownConversion.convertToWord);
            $('#insert-test-Word-doc').click(WordMarkdownConversion.insertWordDoc);
            $('#insert-test-MD-doc').click(WordMarkdownConversion.insertMDDoc);
        });
    };

    WordMarkdownConversion.convertToWord = function () {

        Word.run(function (context) {
            var paragraphs = context.document.body.paragraphs.load('text');

            // The following object enables the code to initialize data objects,
            // that are shared across multiple "then" callbacks within the Word.run,
            // closer to where they are used.
            var sharedDataObjects = {

                /// <field name="paragraphs" type="Word.ParagraphCollection" />
                paragraphs: paragraphs,

                // There are Markdown lists, but no Word lists, at the start,
                // so we can't load them until we've created some. See the
                // convertMDListsToWord method below.
                /// <field name="lists" type="Word.ListCollection" />
                lists: null,

                // Holds groups of paragraphs belonging to MD lists.
                /// <field type="Array" elementType="Word.ParagraphCollection" />
                mdListItems: [],

                // Inline ranges can't be loaded until all paragraph processing is
                // done because the paragraphs that contain these ranges may be recreated.
                // E.g., MD tables are _replaced_ with Word tables, not converted.
                /// <field type="Array" elementType="Word.Range" />
                mdHyperlinkRanges: [],
                /// <field type="Array" elementType="Word.Range" />
                mdBoldRanges: [],
                /// <field type="Array" elementType="Word.Range" />
                candidateMDItalicRanges: []
            };

            return context.sync()
                .then(function () {
                    queueRemovalOfExtraBlankLines(sharedDataObjects.paragraphs);
                    queueConversionOfMDHeadersToWord(sharedDataObjects.paragraphs);
                    queueConversionOfMDCodeBlocksToWord(sharedDataObjects.paragraphs);
                    queueCreationOfWordTables(sharedDataObjects.paragraphs);

                    // Now that the table data is moved into a Word tables, delete the
                    // paragraphs that comprise the original MD tables. But to cope with
                    // a quirk in paragraph deletion the paragraphs collection must be 
                    // reloaded before we can accurately delete any paragraphs.
                    sharedDataObjects.paragraphs = context.document.body.paragraphs.load('text');
                })
                .then(context.sync)
                .then(function () {
                    queueDeletionOfMDTables(sharedDataObjects.paragraphs);
                    queueCreationOfWordLists(sharedDataObjects.paragraphs, sharedDataObjects.mdListItems);

                    // Load the Word lists (which now have just one item each) so we can add items to
                    // them.
                    sharedDataObjects.lists = context.document.body.lists.load('id, paragraphs/text');
                })
                .then(context.sync)
                .then(function () {
                    queueAdditionOfItemsToWordLists(sharedDataObjects.lists, sharedDataObjects.mdListItems);

                    // Cannot load bold or italic text or hyperlink ranges earlier because some
                    // paragraphs in tables (and the inline ranges within them) get recreated when
                    // the Word tables are created.
                    // Find and load all bold and hypertext ranges. (Italic ranges have to wait. See below
                    // for why.)
                    sharedDataObjects.mdHyperlinkRanges = queueFetchOfInlineRanges('hyperlink', context).load('text');
                    sharedDataObjects.mdBoldRanges = queueFetchOfInlineRanges('bold', context).load('text');
                })
                .then(context.sync)
                .then(function () {
                    queueConversionOfMDHyperlinksToWord(sharedDataObjects.mdHyperlinkRanges);
                    queueConversionOfMDBoldStringsToBoldWordRanges(sharedDataObjects.mdBoldRanges);
                    /*
                       Because ranges enclosed in "**" (the MD bold symbol) are a subset of
                       ranges enclosed by "*" (the MD italic symbol), the bold ranges have to
                       be processed -- including removing the enclosing "**" -- and synchronized
                       before the italic ranges can be processed, or even loaded.
                    */
                })
                .then(context.sync)
                .then(function () {
                    sharedDataObjects.candidateMDItalicRanges = queueFetchOfInlineRanges('italic', context).load('text');
                })
                .then(context.sync)
                .then(function () {
                    queueConversionOfMDItalicStringsToItalicWordRanges(sharedDataObjects.candidateMDItalicRanges);
                })
                .then(context.sync);
                // The last "context.sync" is optional, assuming there is at least one earlier
                // "context.sync", since Word.run will do an automatic sync at the end of
                // the batch if there are any pending actions. But it's a best-practice to
                // do an explicit sync.
        })
        .catch(WordMarkdownConversion.errorHandler);

        function queueConversionOfMDHeadersToWord(paragraphs) {
            /// <param name="paragraphs" type="Word.ParagraphCollection" />

            paragraphs.items.forEach(function (paragraph) {

                // In Markdown, all, and only, headers begin with '# ' or '## ' or '### ', etc.
                if (paragraph.text.indexOf('# ') >= 0) {
                    var parsedParagraph = paragraph.text.split('# ');
                    paragraph.insertText(parsedParagraph[1], "Replace");

                    // Note: parsedParagraph[0] does not include the split string '# ', so it
                    // has one fewer '#' character than its heading level.
                    switch (parsedParagraph[0]) {
                        case '':
                            // Use the locale-neutral styleBuiltIn property, so this works
                            // in non-English Word. 
                            paragraph.styleBuiltIn = "Heading1";
                            break;
                        case '#':
                            paragraph.styleBuiltIn = "Heading2";
                            break;
                        case '##':
                            paragraph.styleBuiltIn = "Heading3";
                            break;
                        case '###':
                            paragraph.styleBuiltIn = "Heading4";
                            break;
                        case '####':
                            paragraph.styleBuiltIn = "Heading5";
                            break;
                        case '#####':
                            paragraph.styleBuiltIn = "Heading6";
                            break;
                        default:
                            // Handle corrupted MD headings. (Markdown only supports 6 heading levels.)
                            throw new Error("Invalid Markdown header: \"" + paragraphs.text + "\". Must begin with 1 to 6 '#' characters followed by a space.");
                            break;
                    }
                }
            })
        }

        function queueCreationOfWordTables(paragraphs) {
            /// <param name="paragraphs" type="Word.ParagraphCollection" />

            var previousParagraphIsInTable = false;
            var newWordTable;
            var firstRow;

            paragraphs.items.forEach(function (paragraph) {

                // Assumption: Any paragraph that begins and ends with a "|" symbol
                // is a Markdown table row.
                var mdTableRowParagraph = /\|.*\|/;
                if (mdTableRowParagraph.test(paragraph.text)) {

                    // Every "|" except the last one is the start of a cell/column.
                    var cells = paragraph.text.split('|');

                    // The cells var contains empty items for the starting and ending "|",
                    // but these won't be cells in the Word table, so remove them.
                    var properCells = [];

                    // Loop from the 2nd through the next-to-last.
                    for (var j = 1; j < cells.length - 1; j++) {
                        properCells.push(cells[j]);
                    }

                    if (!previousParagraphIsInTable) {

                        // This is a table row, but the previous paragraph was not,
                        // so we are at the start of a table.
                        previousParagraphIsInTable = true;

                        // Create a Word Table object.
                        newWordTable = paragraph.insertTable(1, properCells.length, "Before");
                        firstRow = newWordTable.rows.getFirst();

                        // Make it resemble standard tables on GitHub.
                        newWordTable.headerRowCount = 1;

                        // In the Table Tools | Design tab of the Word ribbon, each built-in table
                        // style has a name which you can see by hovering the cursor over the
                        // style. Those names (with the spaces removed) are the possible values
                        // of Table.styleBuildIn.
                        newWordTable.styleBuiltIn = "PlainTable1";

                        // Two modifications of PlainTable1 make a better match for GitHub tables.
                        newWordTable.styleBandedRows = false;
                        newWordTable.styleFirstColumn = false;
                    }

                    // Copy the cells from the MD pargraph to the new row,
                    // but skip the Markdown header separator row.
                    var shouldNotSkip = !(paragraph.text.indexOf(':-') >= 0)
                                       &&
                                       !(paragraph.text.indexOf(':=') >= 0);
                    if (shouldNotSkip) {
                        var newRow;
                        if (firstRow) {
                            newRow = firstRow;
                            firstRow = void 0; // Make firstRow undefined again.
                        } else {
                            var rowCollection = newWordTable.addRows("End", 1);
                            newRow = rowCollection.getFirst();
                        }

                        // Writing to row values requires a 2D array.
                        var outerArray = [];
                        outerArray.push(properCells);
                        newRow.values = outerArray;
                    }
                }
                else if (previousParagraphIsInTable) {

                    // This parapraph is not in the table, but the previous one was,
                    // so the table ended.
                    previousParagraphIsInTable = false;
                    newWordTable = null;
                    firstRow = null;
                }
            })
        }

        function queueDeletionOfMDTables(paragraphs) {
            /// <param name="paragraphs" type="Word.ParagraphCollection" />

            // Delete the paragraphs that comprise the original MD tables.
            paragraphs.items.forEach(function (paragraph) {
                var mdTableRowParagraph = /\|.*\|/;
                if (mdTableRowParagraph.test(paragraph.text)) {
                    paragraph.delete();
                }
            })
        }

        function queueConversionOfMDCodeBlocksToWord(paragraphs) {
            /// <param name="paragraphs" type="Word.ParagraphCollection" />

            var inCodeBlock = false;

            paragraphs.items.forEach(function (paragraph) {
                if (paragraph.text === '```') {
                    if (!inCodeBlock) {
                        // This is the start of a MD code block.
                        inCodeBlock = true;
                        // Delete the opening '```' paragraph.
                        paragraph.delete();
                    } else {
                        // This is the end of an MD code block.
                        inCodeBlock = false;
                        // Delete the closing '```' paragraph.
                        paragraph.delete();
                    }
                }
                else if (inCodeBlock) {
                    // Make it somewhat resemble GitHub code blocks.
                    paragraph.styleBuiltIn = 'NoSpacing';
                    paragraph.font.name = 'Consolas';
                    paragraph.font.highlightColor = "LightGray";
                }
            })
        }

        function queueCreationOfWordLists(paragraphs, mdListItems) {
            /// <param name="paragraphs" type="Word.ParagraphCollection" />
            /// <param name="mdListItems" type="Array" elementType="Word.ListCollection" />

            var inList = false;
            var blockOfMDListItems = [];

            var mdListItemParagraph = /^\s*(\*|[0-9]\.)\s.*$/;
            var mdNumberListItemMarker = /[0-9].\s/;

            paragraphs.items.forEach(function (paragraph) {
                if (mdListItemParagraph.test(paragraph.text)) {
                    if (!inList) {

                        // This is the start of a list.
                        inList = true;
                        var list = paragraph.startNewList();

                        // A new list defaults to bullet, so it must be changed if it is
                        // a numbered list. Assumption: A string "n. " (for some n) will appear
                        // in a list item only if it is at the start of a numeric list item.                        
                        if (paragraph.text.search(mdNumberListItemMarker) >= 0) {

                            // A Markdown list that begins at a level other than 0 does not render
                            // properly, so we make the assumption that all lists in the Markdown
                            // begin at level 0.
                            list.setLevelNumbering(0, 'Arabic', [0, "."]);
                            list.setLevelStartingNumber(0, 1);
                            /*
                              No need to remove the original "n. " string (for some n)
                              from the start of the paragraph. Word is intelligent enough
                              to remove that automatically.
                            */
                        }
                        /*
                          No need for an 'else' clause to format a bullet list, because Word
                          automatically formats a bullet list item and even removes the "* "
                          string from the start.
                        */
                    } else {
                        // The paragraph belongs to a list that has already started.
                        // Add it to an array so that it can be added to the list object
                        // with Paragraph.attachToList after the next sync.
                        blockOfMDListItems.push(paragraph);
                    }
                } else {
                    if (inList) {
                        // This paragraph is not in a list, but the preceding one
                        // was, so the list has ended.
                        // Assumption: all lists have at least 2 items.
                        inList = false;
                        mdListItems.push(blockOfMDListItems);
                        blockOfMDListItems = [];
                    }
                }
            })
        }

        function queueAdditionOfItemsToWordLists(lists, mdListItems) {
            /// <param name="lists" type="Word.ListCollection" />
            /// <param name="mdListItems" type="Array" elementType="Word.ListCollection" />

            var mdListLevelTypeMarker = /(\*\s|[0-9].\s)/;
            var mdNumberListItemMarker = /[0-9].\s/;

            // Add addtional items to lists. Each item in mdListItems is an array of paragraphs
            // that represent the 2nd through last items in a list.
            // Each item in the mdListItems array was created immediately after the corresponding
            // Word list was created, so they have matching array indexes; mdListItems[i] has the
            // 2nd through last items that belong to the list lists.items[i].
            for (var i = 0; i < lists.items.length; i++) {
                for (var j = 0; j < mdListItems[i].length; j++) {

                    var parsedItem = mdListItems[i][j].text.split(mdListLevelTypeMarker);

                    // To the left of the "* " or "n. " (for some n), there are four space
                    // characters for every level deep of the list item.
                    var mdLevel = (parsedItem[0].length / 4);

                    mdListItems[i][j].attachToList(lists.items[i].id, mdLevel);
                    mdListItems[i][j].insertText(parsedItem[2], 'Replace')

                    // If the list item is numbered, set its numeric style.
                    if (mdNumberListItemMarker.test(parsedItem[1])) {
                        setNumericStyle(lists.items[i], mdLevel);
                    }
                }
            }
        }

        function setNumericStyle(wordListObject, listLevel) {

            // Rotate among some available numeric styles. The style repeats every
            // 4th level. 
            var level = listLevel % 4;

            var numericStyle;
            switch (level) {
                case 0: // listLevel was 0, 4, or 8.
                    numericStyle = 'Arabic';
                    break;
                case 1:
                    numericStyle = 'UpperLetter';
                    break;
                case 2:
                    numericStyle = 'LowerLetter';
                    break;
                case 3:
                    numericStyle = 'LowerRoman';
                    break;
                default:
                    throw new Error(level + " is not a valid list level. The maximum is 9.");
                    break;
            }
            wordListObject.setLevelNumbering(level, numericStyle, [level, "."]);
            wordListObject.setLevelStartingNumber(level, 1);
        }

        function queueFetchOfInlineRanges(type, context) {
            var searchWildcardExpression;

            // A Word wildcard expressions is like a RegEx, but to escape a character
            // wrap it in "[]" instead of preceding it with "\".
            switch (type) {
                case 'hyperlink':
                    // Markdown hyperlinks have the pattern [link lable](URL).
                    searchWildcardExpression = '[[]*[]][(]*[)]';
                    break;
                case 'bold':
                    // Markdown bold strings have the pattern **bold text here**.
                    searchWildcardExpression = '[*][*]*[*][*]';
                    break;
                case 'italic':
                    // Markdown italic strings have the pattern *italic text here*.
                    searchWildcardExpression = '[*]<*>[*]';
                    break;
                default:
                    throw new Error("Unexpected range type: " + type);
                    break;
            }
            return context.document.body.search(searchWildcardExpression, {
                matchWildCards: true
            });
            /*
               Word search treats a literal "." character as a literal "*" character, so when
               searching for Italic -- [*]<*>[*] --, the preceding search results include ranges
               flanked by "." and ranges flanked by "*", but only the latter are Markdown italic
               ranges. After the next context.sync, the bogus hits must be removed.
            */
        }

        function queueConversionOfMDHyperlinksToWord(mdHyperlinkRanges) {
            /// <param name="mdHyperlinkRanges" type="Array" elementType="Word.Range" />

            mdHyperlinkRanges.items.forEach(function (hyperlinkRange) {
                // The text of each item has the pattern [link lable](URL).
                var parsedHyperlink = hyperlinkRange.text.split('](');

                // The label part becomes the entire string.
                var trimmedLabel = parsedHyperlink[0].split('[')[1];

                // Next line empties the text AND the URL of the hyperlink.
                hyperlinkRange.clear();
                hyperlinkRange.insertText(trimmedLabel, "Start");

                // Process the URL part. This must be done after the label part because the clear()
                // command above blanks the Range.hyperlink property.
                var trimmedURL = parsedHyperlink[1].split(')')[0];
                hyperlinkRange.hyperlink = trimmedURL;
            })
        }

        function queueConversionOfMDBoldStringsToBoldWordRanges(mdBoldRanges) {
            /// <param name="mdBoldRanges" type="Array" elementType="Word.Range" />

            // Remove the MD bold markers ('**') and set the Word bold font.
            mdBoldRanges.items.forEach(function (boldRange) {
                var parsedBoldRange = boldRange.text.split('**');
                var trimmedBoldRange = parsedBoldRange[1];
                boldRange.insertText(trimmedBoldRange, 'Replace');
                boldRange.font.bold = true;
            })
        }

        function queueConversionOfMDItalicStringsToItalicWordRanges(candidateMDItalicRanges) {
            /// <param name="candidateMDItalicRanges" type="Array" elementType="Word.Range" />

            // Create a new array of Range objects that includes only those whose text begins with '*'.
            // See the queueFetchOfInlineRanges method for why this is required.
            var genuineMDItalicRanges = candidateMDItalicRanges.items.filter(function (item) {
                return item.text.indexOf('*') === 0;
            })

            // Remove the MD italic markers ('*') and set the Word italic font.
            genuineMDItalicRanges.forEach(function (italicRange) {
                var parsedItalicRange = italicRange.text.split('*');
                var trimmedItalicRange = parsedItalicRange[1];
                italicRange.insertText(trimmedItalicRange, 'Replace');
                italicRange.font.italic = true;
            })
        }

        function queueRemovalOfExtraBlankLines(paragraphs) {
            /// <param name="paragraphs" type="Word.ParagraphCollection" />

            // Markdown renders two or more blank lines as a single line, but
            // Word renders them all. To get a better approximation in Word of
            // the appearence of the document in rendered Markdown, remove excess
            // blank lines.
            var previousParagraphIsBlankLine = false;

            paragraphs.items.forEach(function (paragraph) {
                if (paragraph.text === '') {
                    if (!previousParagraphIsBlankLine) {
                        // This the first of possibly multiple blank lines.
                        previousParagraphIsBlankLine = true;
                    } else {
                        // This is the second (or later) consecutive blank line,
                        // so queue a command to delete it.
                        paragraph.delete();
                    }
                } else {
                    previousParagraphIsBlankLine = false;
                }
            })
        }
    }

    WordMarkdownConversion.insertWordDoc = function () {
        Word.run(function (context) {
            // WordMarkdownConversion.sampleWordDocument is defined in SampleContent.js
            context.document.body.insertFileFromBase64(WordMarkdownConversion.sampleWordDocument, "Replace");
            return context.sync();
        })
        .catch(WordMarkdownConversion.errorHandler);
    }

    WordMarkdownConversion.insertMDDoc = function () {
        Word.run(function (context) {
            // WordMarkdownConversion.sampleMDDocument is defined in SampleContent.js
            context.document.body.insertText(WordMarkdownConversion.sampleMDDocument, "Replace");
            return context.sync();
        })
        .catch(WordMarkdownConversion.errorHandler);
    }

    WordMarkdownConversion.convertToMarkdown = function () {
        Word.run(function (context) {

            // Load as many objects as possible with as few as possible calls of context.sync.
            // But load only the properties that are actually read by the code.
            var paragraphs = context.document.body.paragraphs.load('tableNestingLevel, font, styleBuiltIn');
            var hyperlinks = context.document.body.getRange().getHyperlinkRanges().load('hyperlink, text');
            var lists = context.document.body.lists.load('levelTypes');
            var tables; // Cannot load tables yet. See below for why.

            // The following object enables the code to initialize data objects,
            // that are shared across multiple "then" callbacks within the Word.run,
            // closer to where they are used.
            var sharedDataObjects = {

                /// <field name="paragraphs" type="Word.ParagraphCollection" />
                paragraphs: paragraphs,

                /// <field name="hyperlinks" type="Word.RangeCollection" />
                hyperlinks: hyperlinks,

                /// <field name="lists" type="Word.ListCollection" />
                lists: lists,

                /// <field name="tables" type="Word.TableCollection" />
                tables: tables,

                // Holds an array of Range collections, one collection per paragraph. Each collection
                // is the words of the paragraph.
                /// <field name="wordRangesInParagraphs" type="Array" elementType="Word.RangeCollection" />
                wordRangesInParagraphs: [],

                // Holds an array of objects (one for each level of each list in the document)
                // that record the level, [sub]list type (Bullet or Number), and child paragraphs
                // at that level of that list.
                levelParagraphs: []
            };
            return context.sync()
                .then(function () {
                    /*
                      Now that paragraphs are loaded, load the word ranges within them
                      that the queueConversionOfBoldAndItalicRanges method is going to process.
                    */
                    sharedDataObjects.paragraphs.items.forEach(function (paragraph) {

                        // Skip headings
                        if (paragraph.styleBuiltIn.indexOf('Heading') === -1) {

                            // Note: wordRanges include "words" with immediately following punctuation
                            // marks. Example: "something," is a word range.
                            var wordRanges = paragraph.getTextRanges([' '], true);

                            // Queue a command to load all the word ranges in the paragraph.
                            wordRanges.load("font/bold, font/italic");
                            sharedDataObjects.wordRangesInParagraphs.push(wordRanges);
                        }
                    })
                    /*
                      Now that the lists are loaded, load the paragraphs at various
                      list levels that the queueConversionOfLists method is going to process.
                    */
                    sharedDataObjects.lists.items.forEach(function (list) {

                        // For all levels of the list. (Word has a 9 level maximum.)
                        for (var j = 0; j < 9; j++) {

                            // Create an object that will record the level and list type
                            // (i.e., Bullet or Number) of the items of the level/list.
                            var paragraphsAtCurrentLevel = {
                                level: j,
                                listTypeAtLevel: list.levelTypes[j]
                            }

                            // Queue commands to add the items (paragraphs) themselves to the object and load them.
                            paragraphsAtCurrentLevel.paragraphCollection = list.getLevelParagraphs(j).load();
                            sharedDataObjects.levelParagraphs.push(paragraphsAtCurrentLevel);
                        }
                    })
                })
                .then(context.sync)
                .then(function () {
                    // Need the file path to handle a quirk in how Word stores bookmark
                    // hyperlinks. See the ensureProperMDLink function.
                    WordMarkdownConversion.fileUrl = getDocumentFilePath();
                })

                // Run the various conversion helper methods. None of these have any
                // internal calls of load() or context.sync() or async calls.
                .then(function () {
                    queueConversionOfBoldAndItalicRanges(sharedDataObjects.wordRangesInParagraphs);
                    queueConversionOfHyperlinkRanges(sharedDataObjects.hyperlinks);
                    queueConversionOfCodeBlocks(sharedDataObjects.paragraphs);
                    queueAdditionOfBlankLineAfterNormalParagraphs(sharedDataObjects.paragraphs);
                    queueConversionOfHeadings(sharedDataObjects.paragraphs);
                    queueConversionOfLists(sharedDataObjects.levelParagraphs);
                })
                .then(function () {
                    // The tables cannot be loaded until now because the queueConversionOfBoldAndItalicRanges()
                    // and queueConversionOfHyperlinkRanges() methods may have changed cell contents in ways
                    // that must be synchronized before the tables are converted.
                    sharedDataObjects.tables = context.document.body.tables.load('values, rowCount');
                })
                .then(context.sync)
                .then(function () {
                    queueConversionOfTables(sharedDataObjects.tables);
                })
                .then(context.sync);
                // The last "context.sync" is optional, assuming there is at least one earlier
                // "context.sync", since Word.run will do an automatic sync at the end of
                // the batch if there are any pending actions. But it's a best-practice to
                // do an explicit sync.
        }).catch(WordMarkdownConversion.errorHandler);


        // Checks each word to see if it's formatted italic or bold, and if so,
        // adds the appropriate Markdown symbols ("*" for italic, "**" for bold).
        function queueConversionOfBoldAndItalicRanges(wordRangesInParagraphs) {
            /// <param name="wordRangesInParagraphs" type="Array" elementType="Word.ParagraphCollection" />

            wordRangesInParagraphs.forEach(function (wordRangesInSingleParagraph) {
                wordRangesInSingleParagraph.items.forEach(function (word, index) {

                    // If several words in a row are styled the same,
                    // the Markdown code is only output at the beginning and end of the string.
                    var previousWord = wordRangesInSingleParagraph.items[index - 1];
                    var nextWord = wordRangesInSingleParagraph.items[index + 1];

                    // Note: word.font.bold is true only when the WHOLE range, including trailing
                    // punctuation, if any, is bold.
                    if (word.font.bold) {
                        if ((typeof previousWord === 'undefined') || !previousWord.font.bold) {
                            word.insertText('**', 'Start');
                        }
                        if ((typeof nextWord === 'undefined') || !nextWord.font.bold) {
                            word.insertText('**', 'End');
                        }
                    }
                    // Note: word.font.italic is true only when the WHOLE range, including trailing
                    // punctuation, is any, is italic.
                    if (word.font.italic) {
                        if ((typeof previousWord === 'undefined') || !previousWord.font.italic) {
                            word.insertText('*', 'Start');
                        }
                        if ((typeof nextWord === 'undefined') || !nextWord.font.italic) {
                            word.insertText('*', 'End');
                        }
                    }
                })
            })
        }

        // Gets a collection of all of the hyperlinks in a document and
        // converts them to Markdown style hyperlinks.
        function queueConversionOfHyperlinkRanges(hyperlinks) {
            /// <param name="hyperlinks" type="Array" elementType="Word.RangeCollection" />

            hyperlinks.items.forEach(function (link) {
                var properLinkURL = ensureProperMDLink(link);
                var mdLink = '[' + link.text + '](' + properLinkURL + ') ';

                // To remove a Word hyperlink, blank the URL.
                link.hyperlink = "";
                link.insertText(mdLink, 'Replace');
            })

            function ensureProperMDLink(hyperlinkRange) {
                var linkURL = hyperlinkRange.hyperlink;

                // If the link.hyperlink property consists of only a bookmark; that is,
                // begins with a "#", Office appends the full local file name of the current
                // document to the front of it. (This will be blank, if the file has never
                // been saved.)
                // The next few lines reverse that because bookmark URLs in Markdown
                // must contain only the bookmark, so that the Markdown renderer is able to
                // append the URL of the online MD file to the front of the bookmark.
                var parsedMDLink = linkURL.split('#');
                if (parsedMDLink[0] === WordMarkdownConversion.fileUrl) {
                    linkURL = '#' + parsedMDLink[1];
                }
                return linkURL;
            }
        }

        // Convert blocks of code paragraphs to Markdown code blocks.
        function queueConversionOfCodeBlocks(paragraphs) {
            /// <param name="paragraphs" type="Word.ParagraphCollection" />

            var previousParagraphIsCode = false;
            var codeBlockParagraphs = [];

            paragraphs.items.forEach(function (paragraph) {

                // Only process a paragraph outside of a table.
                if (paragraph.tableNestingLevel === 0) {

                    // Assumption: Code block paragraphs will use Consolas font.
                    if (paragraph.font.name === 'Consolas') {
                        if (!previousParagraphIsCode) {

                            // This paragraph is Code, but the previous one was not,
                            // so we are at the start of a code block. Add the ``` above the
                            // paragraph to start the Markdown code block.
                            var tripleTickParagraph = paragraph.insertParagraph('```', 'Before');
                            previousParagraphIsCode = true;
                            codeBlockParagraphs.push(tripleTickParagraph);
                        }
                        // Store in order to change its font and background later.
                        codeBlockParagraphs.push(paragraph);
                    }

                    if ((paragraph.font.name != 'Consolas') && (previousParagraphIsCode)) {

                        // This parapraph is not Code, but the previous one was,
                        // so add the Markdown ``` to end the code block.

                        // But Word gives a paragraph that is inserted "Before" the same style.
                        // So change to Normal style temporarily to avoid, for example, giving
                        // the ``` a Header 2 style which would result in Markdown: "## ```".
                        var oldStyle = paragraph.styleBuiltIn;
                        paragraph.styleBuiltIn = 'Normal';
                        paragraph.insertParagraph('```', 'Before');
                        paragraph.styleBuiltIn = oldStyle;
                        previousParagraphIsCode = false;
                    }
                }
            })

            // Change font and background color of code block paragraphs to make
            // look more like plain Markdown text.
            codeBlockParagraphs.forEach(function (codeParagraph) {
                codeParagraph.font.name = 'Calibri (Body)';
                codeParagraph.font.highlightColor = null;
            })
        }

        // Formatting of Markdown text that follows an ordinary paragraph requires
        // a blank line after the paragraph, but Word paragraphs with Normal style
        // typically don't have one. Multiple blank lines are treated as one in Markdown,
        // So there's no harm if a blank line is added where there already is one.
        function queueAdditionOfBlankLineAfterNormalParagraphs(paragraphs) {
            /// <param name="paragraphs" type="Word.ParagraphCollection" />

            paragraphs.items.forEach(function (paragraph) {

                // Only process a paragraph outside of a table.
                if (paragraph.tableNestingLevel === 0) {
                    if (paragraph.styleBuiltIn === 'Normal') {
                        paragraph.styleBuiltIn = 'Normal';
                        paragraph.insertParagraph('', 'After');
                    }
                }
            })
        }

        // Finds all the headers in the document and converts them to Markdown.
        function queueConversionOfHeadings(paragraphs) {
            /// <param name="paragraphs" type="Word.ParagraphCollection" />

            paragraphs.items.forEach(function (paragraph) {

                // Only headers have a styleBuiltIn value with "Heading" in the name.
                if (paragraph.styleBuiltIn.indexOf('Heading') === 0) {

                    switch (paragraph.styleBuiltIn) {
                        case 'Heading1':
                            paragraph.insertText('# ', 'Start');
                            break;
                        case 'Heading2':
                            paragraph.insertText('## ', 'Start');
                            break;
                        case 'Heading3':
                            paragraph.insertText('### ', 'Start');
                            break;
                        case 'Heading4':
                            paragraph.insertText('#### ', 'Start');
                            break;
                        case 'Heading5':
                            paragraph.insertText('##### ', 'Start');
                            break;
                        case 'Heading6':
                            paragraph.insertText('###### ', 'Start');
                            break;
                        default:
                            // Markdown supports only 6 levels of headings. But instead of throwing an
                            // error, just let a paragraph with a "Heading" style greater than 6 be 
                            // changed to a non-heading in Markdown by the first line after the switch block.
                            break;
                    }
                    // Turn off Word heading formatting so it looks more like plain Markdown.
                    paragraph.styleBuiltIn = 'Normal';

                    // Some Markdown renderers fail to correctly render lists or tables if the
                    // immediately preceding, or following, paragraph is a header, but Word paragraphs
                    // with a Header style typically are not followed/preceded by a blank line,
                    // so add one. Multiple blank lines are treated as one in Markdown, so there's
                    // no harm if a blank line is added where there already is one.
                    paragraph.insertParagraph('', 'Before');
                    paragraph.insertParagraph('', 'After');
                }
            })
        }

        // Convert Word lists into Markdown syntax.
        function queueConversionOfLists(levelParagraphs) {
            /// <param name="levelParagraphs" type="Word.ParagraphCollection" />

            // Note: levelParagraphs holds an array of objects (one for each level of each list in the document)
            // that record the level, [sub]list type [Bullet or Number], and child paragraphs
            // at that level of that list.
            levelParagraphs.forEach(function (levelDefinition) {

                // Note: levelDefinition.paragraphCollection is a Word.ParagraphCollection
                // levelDefinition.paragraphCollection.items is an array of Word.Paragraph objects.
                levelDefinition.paragraphCollection.items.forEach(function (wordParagraph) {

                    var mdListItemPrefix = getMarkdownListItemPrefix(levelDefinition.listTypeAtLevel,
                                                                     levelDefinition.level);

                    // Insert Markdown list item (level-relative and type-relative) prefix.
                    wordParagraph.insertText(mdListItemPrefix, 'Start');

                    // Turn off Word list style, so the list resembles Markdown text.
                    wordParagraph.styleBuiltIn = 'Normal';
                })
            })

            // Returns a string of '* ' or '1. ' symbol, preceded by 4 spaces
            // for each level deep in a (possibly nested) list.
            function getMarkdownListItemPrefix(wordListType, listLevel) {

                var listTypeSymbol = getMarkdownSymbol(wordListType);
                var itemPrefix = '';
                for (var i = 0; i < listLevel; i++) {
                    itemPrefix = itemPrefix + "    "; // 4 spaces
                }
                return itemPrefix + listTypeSymbol + ' ';

                function getMarkdownSymbol(wordListType) {
                    switch (wordListType) {
                        case 'Bullet':
                            return '*';
                        case 'Number':
                            return '1.';
                        case 'Picture':
                            throw new Error("Cannot convert Word picture list type to Markdown.");
                            break;
                        default:
                            throw new Error("Unknown Word list type. Should be 'Number', 'Bullet', or 'Picture'.");
                            break;
                    }
                }
            }
        }

        // Gets a collection of all of the tables in the document and converts
        // them to Markdown-style tables.
        function queueConversionOfTables(tables) {
            /// <param name="tables" type="Word.TableCollection" />

            tables.items.forEach(function (table) {
                table.values.forEach(function (cellValues, index) {

                    // Create a Markdown table above the Word table. Add each
                    // row below the preceding row, which means just before the
                    // Word table.
                    var rowParagraph = table.insertParagraph('| ', 'Before');
                    rowParagraph.styleBuiltIn = 'Normal';

                    // Copy each cell from the Word table to the Markdown table.
                    cellValues.forEach(function (cellValue) {
                        rowParagraph.insertText(cellValue + ' | ', 'End');
                    })

                    // If the row that was just created is the first one, then insert the
                    // Markdown separator row. (If your Markdown renderer supports rows
                    // without headers -- GitHub and others do NOT -- then you could check
                    // (table.heardRowCount > 0) to see if the Word table has a header.)
                    if (index === 0) {
                        var mdSeparatorRow = '|';
                        const mdSeparatorCell = ':--|';

                        // "for" instead of "forEach" because its only the length of cellValues
                        // that matters. The loop does not read or write cellValues.
                        for (var k = 0; k < cellValues.length; k++) {
                            mdSeparatorRow = mdSeparatorRow + mdSeparatorCell;
                        }
                        table.insertParagraph(mdSeparatorRow, 'Before');
                    }
                })
                // Remove the original Word table.
                table.delete();
            })
        }
    }

    // Gets the full path and file name of the Word document.
    function getDocumentFilePath() {

        // Wrap call to Office 2013 asynchronous method in a promise to be consistent with
        // the Office 2016 asynchronous architecture.
        return new OfficeExtension.Promise(function (resolve, reject) {
            try {
                // Need to use the Office 2013 JavaScript APIs to get the full file path and name.
                // This data is needed to handle a quirk in how Word stores bookmark hyperlinks. See
                // the method ensureProperMDLink in this file. Note: the Office.context.url
                // property returns an HTTP URL, not the file path and name that is needed in this case.
                Office.context.document.getFilePropertiesAsync(function (asyncResult) {
                    resolve(asyncResult.value.url);
                });
            }
            catch (error) {
                reject(WordMarkdownConversion.errorHandler(error));
            }
        })
    }

    WordMarkdownConversion.errorHandler = function (error) {
        WordMarkdownConversion.showNotification(error);
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
            console.log("Trace info: " + JSON.stringify(error.traceMessages));
        }
    }

    // Helper function for displaying notifications in a message banner.
    WordMarkdownConversion.showNotification = function (content) {
        $("#notificationBody").text(content);
        FabricComponents.messageBanner.showBanner();
        FabricComponents.messageBanner.toggleExpansion();
    }

    window.WordMarkdownConversion = WordMarkdownConversion;
})();

