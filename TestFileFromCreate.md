
# Body Object (JavaScript API for Word)

Represents the body of a document or a section.

_Applies to: Office 2016_

| Property | Type | Description |
| --- | --- | --- |
| style | string | Gets or sets the style used for the body. This is the name of the pre-installed or custom style. |
| text | string | Gets the text of the body. Use the insertText method to insert text. Read-only. |

_See property access [examples.](#property-access-examples)_

## Relationships

| Relationship | Type | Description |
| --- | --- | --- |
| contentControls | [ContentControlCollection](contentcontrolcollection.md) | Gets the collection of rich text content control objects that are in the body. Read-only. |
| font | [Font](font.md) | Gets the text format of the body. Use this to get and set font name, size, color, and other properties. Read-only. |
| inlinePictures | [InlinePictureCollection](inlinepicturecollection.md) | Gets the collection of inlinePicture objects that are in the body. The collection does not include floating images. Read-only. |
| paragraphs | [ParagraphCollection](paragraphcollection.md) | Gets the collection of paragraph objects that are in the body. Read-only. |
| parentContentControl | [ContentControl](contentcontrol.md) | Gets the content control that contains the body. Returns null if there isn't a parent content control. Read-only. |

## Methods

| Method | Return Type | Description |
| --- | --- | --- |
| [clear()](#clear) | void | Clears the contents of the body object. The user can perform the undo operation on the cleared content. |
| [getHtml()](#gethtml) | string | Gets the HTML representation of the body object. |
| [getOoxml()](#getooxml) | string | Gets the OOXML (Office Open XML) representation of the body object. |
| [insertBreak(breakType: BreakType, insertLocation: InsertLocation)](#insertbreakbreaktype-breaktype-insertlo) | void | Inserts a break at the specified location. The insertLocation value can be 'Start' or 'End'. |
| [insertContentControl()](#insertcontentcontrol) | [ContentControl](contentcontrol.md) | Wraps the body object with a Rich Text content control. |
| [insertFileFromBase64(base64File: string, insertLocation: InsertLocation)](#insertfilefrombase64base64file-string-i) | [Range](range.md) | Inserts a document into the body at the specified location. The insertLocation value can be 'Replace', 'Start' or 'End'. |
| [insertHtml(html: string, insertLocation: InsertLocation)](#inserthtmlhtml-string-insertlocation-in) | [Range](range.md) | Inserts HTML at the specified location. The insertLocation value can be 'Replace', 'Start' or 'End'. |
| [insertOoxml(ooxml: string, insertLocation: InsertLocation)](#insertooxmlooxml-string-insertlocation-) | [Range](range.md) | Inserts OOXML at the specified location. The insertLocation value can be 'Replace', 'Start' or 'End'. |
| [insertParagraph(paragraphText: string, insertLocation: InsertLocation)](#insertparagraphparagraphtext-string-ins) | [Paragraph](paragraph.md) | Inserts a paragraph at the specified location. The insertLocation value can be 'Start' or 'End'. |
| [insertText(text: string, insertLocation: InsertLocation)](#inserttexttext-string-insertlocation-in) | [Range](range.md) | Inserts text into the body at the specified location. The insertLocation value can be 'Replace', 'Start' or 'End'. |
| [load(param: object)](#loadparam-object) | void | Fills the proxy object created in JavaScript layer with property and object values specified in the parameter. |
| [search(searchText: string, searchOptions: ParamTypeStrings.SearchOptions)](#searchsearchtext-string-searchoptions-p) | [SearchResultCollection](searchresultcollection.md) | Performs a search with the specified searchOptions on the scope of the body object. The search results are a collection of range objects. |
| [select()](#select) | void | Selects the body and navigates the Word UI to it. |

## Method Details

### clear()

Clears the contents of the body object. The user can perform the undo operation on the cleared content.

#### Syntax

<pre>bodyObject.clear();</pre>

#### Parameters

None

#### Returns

void

#### Examples

<pre>// Run a batch operation against the Word object model.
Word.run(function (context) {
    // Create a proxy object for the document body.
    var body = context.document.body;

    // Queue a commmand to clear the contents of the body.
    body.clear();

    // Synchronize the document state by executing the queued-up commands, 
    // and return a promise to indicate task completion.
    return context.sync().then(function () {
        console.log('Cleared the body contents.');

    });

})
.catch(function (error) {
    console.log("Error: "

+ JSON.stringify(error));

    if (error instanceof OfficeExtension.Error) {
        console.log("Debug info: "

+ JSON.stringify(error.debugInfo));

    }
});</pre>

### getHtml()

Gets the HTML representation of the body object.

#### Syntax

<pre>bodyObject.getHtml();</pre>

#### Parameters

None

#### Returns

string

#### Examples

<pre>// Run a batch operation against the Word object model.
Word.run(function (context) {
    // Create a proxy object for the document body.
    var body = context.document.body;

    // Queue a commmand to get the HTML contents of the body.
    var bodyHTML = body.getHtml();

    // Synchronize the document state by executing the queued-up commands, 
    // and return a promise to indicate task completion.
    return context.sync().then(function () {
        console.log("Body HTML contents: "

+ bodyHTML.value);

    });

})
.catch(function (error) {
    console.log("Error: "

+ JSON.stringify(error));

    if (error instanceof OfficeExtension.Error) {
        console.log("Debug info: "

+ JSON.stringify(error.debugInfo));

    }
});</pre>

### getOoxml()

Gets the OOXML (Office Open XML) representation of the body object.

#### Syntax

<pre>bodyObject.getOoxml();</pre>

#### Parameters

None

#### Returns

string

#### Examples

<pre>// Run a batch operation against the Word object model.
Word.run(function (context) {
    // Create a proxy object for the document body.
    var body = context.document.body;

    // Queue a commmand to get the OOXML contents of the body.
    var bodyOOXML = body.getOoxml();

    // Synchronize the document state by executing the queued-up commands, 
    // and return a promise to indicate task completion.
    return context.sync().then(function () {
        console.log("Body OOXML contents: "

+ bodyOOXML.value);

    });

})
.catch(function (error) {
    console.log("Error: "

+ JSON.stringify(error));

    if (error instanceof OfficeExtension.Error) {
        console.log("Debug info: "

+ JSON.stringify(error.debugInfo));

    }
});</pre>

### insertBreak(breakType: BreakType, insertLocation: InsertLocation)

Inserts a break at the specified location. The insertLocation value can be 'Start' or 'End'.

#### Syntax

<pre>bodyObject.insertBreak(breakType, insertLocation);</pre>

#### Parameters

| Parameter | Type | Description |
| --- | --- | --- |
| breakType | BreakType | Required. The break type to add to the body. |
| insertLocation | InsertLocation | Required. The value can be 'Start' or 'End'. |

#### Returns

void

#### Examples

<pre>// Run a batch operation against the Word object model.
Word.run(function (ctx) {
    // Create a proxy object for the document body.
    var body = ctx.document.body;

    // Queue a commmand to insert a page break at the start of the document body.
    body.insertBreak(Word.BreakType.page, Word.InsertLocation.start);

    // Synchronize the document state by executing the queued-up commands, 
    // and return a promise to indicate task completion.
    return ctx.sync().then(function () {
        console.log('Added a page break at the start of the document body.');

    });

})
.catch(function (error) {
    console.log("Error: "

+ JSON.stringify(error));

    if (error instanceof OfficeExtension.Error) {
        console.log("Debug info: "

+ JSON.stringify(error.debugInfo));

    }
});</pre>

### insertContentControl()

Wraps the body object with a Rich Text content control.

#### Syntax

<pre>bodyObject.insertContentControl();</pre>

#### Parameters

None

#### Returns

[ContentControl](contentcontrol.md)

#### Examples

<pre>// Run a batch operation against the Word object model.
Word.run(function (context) {
    // Create a proxy object for the document body.
    var body = context.document.body;

    // Queue a commmand to wrap the body in a content control.
    body.insertContentControl();

    // Synchronize the document state by executing the queued-up commands, 
    // and return a promise to indicate task completion.
    return context.sync().then(function () {
        console.log('Wrapped the body in a content control.');

    });

})
.catch(function (error) {
    console.log('Error: ' + JSON.stringify(error));

    if (error instanceof OfficeExtension.Error) {
        console.log('Debug info: ' + JSON.stringify(error.debugInfo));

    }
});</pre>

### insertFileFromBase64(base64File: string, insertLocation: InsertLocation)

Inserts a document into the body at the specified location. The insertLocation value can be 'Replace', 'Start' or 'End'.

#### Syntax

<pre>bodyObject.insertFileFromBase64(base64File, insertLocation);</pre>

#### Parameters

| Parameter | Type | Description |
| --- | --- | --- |
| base64File | string | Required. The base64 encoded file contents to be inserted. |
| insertLocation | InsertLocation | Required. The value can be 'Replace', 'Start' or 'End'. |

#### Returns

[Range](range.md)

#### Examples

<pre>// Run a batch operation against the Word object model.
Word.run(function (context) {
    // Create a proxy object for the document body.
    var body = context.document.body;

    // Queue a commmand to insert base64 encoded .docx at the beginning of the content body.
    // You will need to implement getBase64() to pass in a string of a base64 encoded docx file.
    body.insertFileFromBase64(getBase64(), Word.InsertLocation.start);

    // Synchronize the document state by executing the queued-up commands, 
    // and return a promise to indicate task completion.
    return context.sync().then(function () {
        console.log('Added base64 encoded text to the beginning of the document body.');

    });

})
.catch(function (error) {
    console.log('Error: ' + JSON.stringify(error));

    if (error instanceof OfficeExtension.Error) {
        console.log('Debug info: ' + JSON.stringify(error.debugInfo));

    }
});</pre>

### insertHtml(html: string, insertLocation: InsertLocation)

Inserts HTML at the specified location. The insertLocation value can be 'Replace', 'Start' or 'End'.

#### Syntax

<pre>bodyObject.insertHtml(html, insertLocation);</pre>

#### Parameters

| Parameter | Type | Description |
| --- | --- | --- |
| html | string | Required. The HTML to be inserted in the document. |
| insertLocation | InsertLocation | Required. The value can be 'Replace', 'Start' or 'End'. |

#### Returns

[Range](range.md)

#### Examples

<pre>// Run a batch operation against the Word object model.
Word.run(function (context) {
    // Create a proxy object for the document body.
    var body = context.document.body;

    // Queue a commmand to insert HTML in to the beginning of the body.
    body.insertHtml('<strong>This is text inserted with body.insertHtml()</strong>', Word.InsertLocation.start);

    // Synchronize the document state by executing the queued-up commands, 
    // and return a promise to indicate task completion.
    return context.sync().then(function () {
        console.log('HTML added to the beginning of the document body.');

    });

})
.catch(function (error) {
    console.log('Error: ' + JSON.stringify(error));

    if (error instanceof OfficeExtension.Error) {
        console.log('Debug info: ' + JSON.stringify(error.debugInfo));

    }
});</pre>

### insertOoxml(ooxml: string, insertLocation: InsertLocation)

Inserts OOXML at the specified location. The insertLocation value can be 'Replace', 'Start' or 'End'.

#### Syntax

<pre>bodyObject.insertOoxml(ooxml, insertLocation);</pre>

#### Parameters

| Parameter | Type | Description |
| --- | --- | --- |
| ooxml | string | Required. The OOXML to be inserted. |
| insertLocation | InsertLocation | Required. The value can be 'Replace', 'Start' or 'End'. |

#### Returns

[Range](range.md)

#### Examples

<pre>// Run a batch operation against the Word object model.
Word.run(function (context) {
    // Create a proxy object for the document body.
    var body = context.document.body;

    // Queue a commmand to insert OOXML in to the beginning of the body.
    body.insertOoxml("<w:p xmlns:w='http://schemas.microsoft.com/office/word/2003/wordml'><w:r><w:rPr><w:b/><w:b-cs/>"

+
                     "<w:color w:val='FF0000'/><w:sz w:val='28'/><w:sz-cs w:val='28'/></w:rPr><w:t>Hello world (this"

+
                     "

should be bold, red, size 14).</w:t></w:r></w:p>", Word.InsertLocation.start);

    // Synchronize the document state by executing the queued-up commands, 
    // and return a promise to indicate task completion.
    return context.sync().then(function () {
        console.log('OOXML added to the beginning of the document body.');

    });

})
.catch(function (error) {
    console.log('Error: ' + JSON.stringify(error));

    if (error instanceof OfficeExtension.Error) {
        console.log('Debug info: ' + JSON.stringify(error.debugInfo));

    }
});</pre>

### insertParagraph(paragraphText: string, insertLocation: InsertLocation)

Inserts a paragraph at the specified location. The insertLocation value can be 'Start' or 'End'.

#### Syntax

<pre>bodyObject.insertParagraph(paragraphText, insertLocation);</pre>

#### Parameters

| Parameter | Type | Description |
| --- | --- | --- |
| paragraphText | string | Required. The paragraph text to be inserted. |
| insertLocation | InsertLocation | Required. The value can be 'Start' or 'End'. |

#### Returns

[Paragraph](paragraph.md)

#### Examples

<pre>// Run a batch operation against the Word object model.
Word.run(function (context) {
    // Create a proxy object for the document body.
    var body = context.document.body;

    // Queue a commmand to insert the paragraph at the end of the document body.
    body.insertParagraph('Content of a new paragraph', Word.InsertLocation.end);

    // Synchronize the document state by executing the queued-up commands, 
    // and return a promise to indicate task completion.
    return context.sync().then(function () {
        console.log('Paragraph added at the end of the document body.');

    });

})
.catch(function (error) {
    console.log('Error: ' + JSON.stringify(error));

    if (error instanceof OfficeExtension.Error) {
        console.log('Debug info: ' + JSON.stringify(error.debugInfo));

    }
});</pre>

### insertText(text: string, insertLocation: InsertLocation)

Inserts text into the body at the specified location. The insertLocation value can be 'Replace', 'Start' or 'End'.

#### Syntax

<pre>bodyObject.insertText(text, insertLocation);</pre>

#### Parameters

| Parameter | Type | Description |
| --- | --- | --- |
| text | string | Required. Text to be inserted. |
| insertLocation | InsertLocation | Required. The value can be 'Replace', 'Start' or 'End'. |

#### Returns

[Range](range.md)

#### Examples

<pre>// Run a batch operation against the Word object model.
Word.run(function (context) {
    // Create a proxy object for the document body.
    var body = context.document.body;

    // Queue a commmand to insert text in to the beginning of the body.
    body.insertText('This is text inserted with body.insertText()', Word.InsertLocation.start);

    // Synchronize the document state by executing the queued-up commands, 
    // and return a promise to indicate task completion.
    return context.sync().then(function () {
        console.log('Text added to the beginning of the document body.');

    });

})
.catch(function (error) {
    console.log('Error: ' + JSON.stringify(error));

    if (error instanceof OfficeExtension.Error) {
        console.log('Debug info: ' + JSON.stringify(error.debugInfo));

    }
});</pre>

### load(param: object)

Fills the proxy object created in JavaScript layer with property and object values specified in the parameter.

#### Syntax

<pre>object.load(param);</pre>

#### Parameters

| Parameter | Type | Description |
| --- | --- | --- |
| param | object | Optional. Accepts parameter and relationship names as delimited string or an array. Or, provide [loadOption](loadoption.md) object. |

#### Returns

void

#### Examples

<pre>// Run a batch operation against the Word object model.
Word.run(function (context) {
    // Create a proxy object for the document body.
    var body = context.document.body;

    // Queue a commmand to load font and style information for the document body.
    context.load(body, 'font/size, font/name, font/color, style');

    // Synchronize the document state by executing the queued-up commands, 
    // and return a promise to indicate task completion.
    return context.sync().then(function () {
        // Show the results of the load method. Here we show the
        // property values on the body object.
        var results = 'Font size: ' + body.font.size +
                      ';

Font name: ' + body.font.name +
                      ';

Font color: ' + body.font.color +
                      ';

Body style: ' + body.style;

        console.log(results);

    });

})
.catch(function (error) {
    console.log('Error: ' + JSON.stringify(error));

    if (error instanceof OfficeExtension.Error) {
        console.log('Debug info: ' + JSON.stringify(error.debugInfo));

    }
});</pre>

### search(searchText: string, searchOptions: ParamTypeStrings.SearchOptions)

Performs a search with the specified searchOptions on the scope of the body object. The search results are a collection of range objects.

#### Syntax

<pre>bodyObject.search(searchText, searchOptions);</pre>

#### Parameters

| Parameter | Type | Description |
| --- | --- | --- |
| searchText | string | Required. The search text. |
| searchOptions | ParamTypeStrings.SearchOptions | Optional. Optional. Options for the search. |

#### Returns

[SearchResultCollection](searchresultcollection.md)

#### Examples

<pre>// Run a batch operation against the Word object model.
Word.run(function (context) {
    // Create a proxy object for the document body.
    var body = context.document.body;

    // Setup the search options.
    var options = Word.SearchOptions.newObject(context);

    options.matchCase = false
    // Queue a commmand to search the document.
    var searchResults = context.document.body.search('video', options);

    // Queue a commmand to load the results.
    context.load(searchResults, 'text, font');

    // Synchronize the document state by executing the queued-up commands, 
    // and return a promise to indicate task completion.
    return context.sync().then(function () {
        var results = 'Found count: ' + searchResults.items.length + 
                      ';

we highlighted the results.';

        // Queue a command to change the font for each found item. 
        for (var i = 0;

i <

searchResults.items.length;

i++) {
          searchResults.items[i].font.color = '#FF0000'    // Change color to Red
          searchResults.items[i].font.highlightColor = '#FFFF00';

          searchResults.items[i].font.bold = true;

        }
        // Synchronize the document state by executing the queued-up commands, 
        // and return a promise to indicate task completion.
        return context.sync().then(function () {
            console.log(results);

        });

    });

})
.catch(function (error) {
    console.log('Error: ' + JSON.stringify(error));

    if (error instanceof OfficeExtension.Error) {
        console.log('Debug info: ' + JSON.stringify(error.debugInfo));

    }
});</pre>

### select()

Selects the body and navigates the Word UI to it.

#### Syntax

<pre>bodyObject.select();</pre>

#### Parameters

None

#### Returns

void

#### Examples

<pre>// Run a batch operation against the Word object model.
Word.run(function (context) {
    // Create a proxy object for the document body.
    var body = context.document.body;

    // Queue a commmand to select the document body. The Word UI will 
    // move to the selected document body.
    body.select();

    // Synchronize the document state by executing the queued-up commands, 
    // and return a promise to indicate task completion.
    return context.sync().then(function () {
        console.log('Selected the document body.');

    });

})
.catch(function (error) {
    console.log('Error: ' + JSON.stringify(error));

    if (error instanceof OfficeExtension.Error) {
        console.log('Debug info: ' + JSON.stringify(error.debugInfo));

    }
});</pre>

## Property access examples

### Get the text property on the body object

<pre>// Run a batch operation against the Word object model.
Word.run(function (context) {
    // Create a proxy object for the document body.
    var body = context.document.body;

    // Queue a commmand to load the text in document body.
    context.load(body, 'text');

    // Synchronize the document state by executing the queued-up commands, 
    // and return a promise to indicate task completion.
    return context.sync().then(function () {
        console.log("Body contents: "

+ body.text);

    });

})
.catch(function (error) {
    console.log("Error: "

+ JSON.stringify(error));

    if (error instanceof OfficeExtension.Error) {
        console.log("Debug info: "

+ JSON.stringify(error.debugInfo));

    }
});</pre>

### Get the style and the font size, font name, and font color properties on the body object.

<pre>// Run a batch operation against the Word object model.
Word.run(function (context) {
    // Create a proxy object for the document body.
    var body = context.document.body;

    // Queue a commmand to load font and style information for the document body.
    context.load(body, 'font/size, font/name, font/color, style');

    // Synchronize the document state by executing the queued-up commands, 
    // and return a promise to indicate task completion.
    return context.sync().then(function () {
        // Show the results of the load method. Here we show the
        // property values on the body object.
        var results = 'Font size: ' + body.font.size +
                      ';

Font name: ' + body.font.name +
                      ';

Font color: ' + body.font.color +
                      ';

Body style: ' + body.style;

        console.log(results);

    });

})
.catch(function (error) {
    console.log('Error: ' + JSON.stringify(error));

    if (error instanceof OfficeExtension.Error) {
        console.log('Debug info: ' + JSON.stringify(error.debugInfo));

    }</pre>

});
