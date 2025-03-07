# Manage body

Shows how to manage the document body.

This sample shows how to manage the content of the document body.

```typescript
async function getFontProperties() {
  // Gets the style and the font size, font name, and font color properties on the body object.
  // Run a batch operation against the Word object model.
  await Word.run(async (context) => {
    // Create a proxy object for the document body.
    const body: Word.Body = context.document.body;

    // Queue a command to load font and style information for the document body.
    body.load("font/size, font/name, font/color, style");

    // Synchronize the document state by executing the queued commands, and return a promise to indicate task completion.
    await context.sync();

    // Show font-related property values on the body object.
    const results =
      "Font size: " +
      body.font.size +
      "; Font name: " +
      body.font.name +
      "; Font color: " +
      body.font.color +
      "; Body style: " +
      body.style;

    console.log(results);
  });
}

async function getHTML() {
  // Gets the HTML that represents the content of the body.
  // Run a batch operation against the Word object model.
  await Word.run(async (context) => {
    // Create a proxy object for the document body.
    const body: Word.Body = context.document.body;

    // Queue a command to get the HTML contents of the body.
    const bodyHTML = body.getHtml();

    // Synchronize the document state by executing the queued commands, and return a promise to indicate task completion.
    await context.sync();

    console.log("Body contents (HTML): " + bodyHTML.value);
  });
}

async function getOOXML() {
  // Gets the OOXML that represents the content of the body.
  // Run a batch operation against the Word object model.
  await Word.run(async (context) => {
    // Create a proxy object for the document body.
    const body: Word.Body = context.document.body;

    // Queue a command to get the OOXML contents of the body.
    const bodyOOXML = body.getOoxml();

    // Synchronize the document state by executing the queued commands, and return a promise to indicate task completion.
    await context.sync();

    console.log("Body contents (OOXML): " + bodyOOXML.value);
  });
}

async function getText() {
  // Gets the text content of the body.
  // Run a batch operation against the Word object model.
  await Word.run(async (context) => {
    // Create a proxy object for the document body.
    const body: Word.Body = context.document.body;

    // Queue a command to load the text in document body.
    body.load("text");

    // Synchronize the document state by executing the queued commands, and return a promise to indicate task completion.
    await context.sync();

    console.log("Body contents (text): " + body.text);
  });
}

async function insertContentControl() {
  // Creates a content control using the document body.
  // Run a batch operation against the Word object model.
  await Word.run(async (context) => {
    // Create a proxy object for the document body.
    const body: Word.Body = context.document.body;

    // Queue a command to wrap the body in a content control.
    body.insertContentControl();

    // Synchronize the document state by executing the queued commands, and return a promise to indicate task completion.
    await context.sync();

    console.log("Wrapped the body in a content control.");
  });
}

async function insertPageBreak() {
  // Inserts a page break at the beginning of the document.
  // Run a batch operation against the Word object model.
  await Word.run(async (context) => {
    // Create a proxy object for the document body.
    const body: Word.Body = context.document.body;

    // Queue a command to insert a page break at the start of the document body.
    body.insertBreak(Word.BreakType.page, Word.InsertLocation.start);

    // Synchronize the document state by executing the queued commands, and return a promise to indicate task completion.
    await context.sync();

    console.log("Added a page break at the start of the document body.");
  });
}

let externalDocument;

function getBase64() {
  // Retrieve the file and set up an HTML FileReader element.
  const myFile = <HTMLInputElement>document.getElementById("file");
  const reader = new FileReader();

  reader.onload = (event) => {
    // Remove the metadata before the Base64-encoded string.
    const startIndex = reader.result.toString().indexOf("base64,");
    externalDocument = reader.result.toString().substr(startIndex + 7);
  };

  // Read the file as a data URL so that we can parse the Base64-encoded string.
  reader.readAsDataURL(myFile.files[0]);
}

async function insertExternalBody() {
  // Inserts the body from the external document at the beginning of this document.
  // Run a batch operation against the Word object model.
  await Word.run(async (context) => {
    // Create a proxy object for the document body.
    const body: Word.Body = context.document.body;

    // Queue a command to insert the Base64-encoded string representation of the body of the selected .docx file at the beginning of the current document.
    body.insertFileFromBase64(externalDocument, Word.InsertLocation.start);

    // Synchronize the document state by executing the queued commands, and return a promise to indicate task completion.
    await context.sync();

    console.log("Added Base64-encoded text to the beginning of the document body.");
  });
}

async function insertHTML() {
  // Inserts the HTML at the beginning of this document.
  // Run a batch operation against the Word object model.
  await Word.run(async (context) => {
    // Create a proxy object for the document body.
    const body: Word.Body = context.document.body;

    // Queue a command to insert HTML at the beginning of the document.
    body.insertHtml("<strong>This is text inserted with body.insertHtml()</strong>", Word.InsertLocation.start);

    // Synchronize the document state by executing the queued commands, and return a promise to indicate task completion.
    await context.sync();

    console.log("HTML added to the beginning of the document body.");
  });
}

async function insertImageInline() {
  // Inserts an image inline at the beginning of this document.
  // Run a batch operation against the Word object model.
  await Word.run(async (context) => {
    // Create a proxy object for the document body.
    const body: Word.Body = context.document.body;

    // Base64-encoded image to insert inline.
    const base64EncodedImg =
      "iVBORw0KGgoAAAANSUhEUgAAAB4AAAANCAIAAAAxEEnAAAAAAXNSR0IArs4c6QAAAARnQU1BAACxjwv8YQUAAAAJcEhZcwAADsMAAA7DAcdvqGQAAACFSURBVDhPtY1BEoQwDMP6/0+XgIMTBAeYoTqso9Rkx1zG+tNj1H94jgGzeNSjteO5vtQQuG2seO0av8LzGbe3anzRoJ4ybm/VeKEerAEbAUpW4aWQCmrGFWykRzGBCnYy2ha3oAIq2MloW9yCCqhgJ6NtcQsqoIKdjLbFLaiACnYyf2fODbrjZcXfr2F4AAAAAElFTkSuQmCC";

    // Queue a command to insert a Base64-encoded image at the beginning of the current document.
    body.insertInlinePictureFromBase64(base64EncodedImg, Word.InsertLocation.start);

    // Synchronize the document state by executing the queued commands, and return a promise to indicate task completion.
    await context.sync();

    console.log("Added a Base64-encoded image to the beginning of the document body.");
  });
}

async function insertOOXML() {
  // Inserts OOXML at the beginning of this document.
  // Run a batch operation against the Word object model.
  await Word.run(async (context) => {
    // Create a proxy object for the document body.
    const body: Word.Body = context.document.body;

    // Queue a command to insert OOXML at the beginning of the body.
    body.insertOoxml(
      "<pkg:package xmlns:pkg='http://schemas.microsoft.com/office/2006/xmlPackage'><pkg:part pkg:name='/_rels/.rels' pkg:contentType='application/vnd.openxmlformats-package.relationships+xml' pkg:padding='512'><pkg:xmlData><Relationships xmlns='http://schemas.openxmlformats.org/package/2006/relationships'><Relationship Id='rId1' Type='http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument' Target='word/document.xml'/></Relationships></pkg:xmlData></pkg:part><pkg:part pkg:name='/word/document.xml' pkg:contentType='application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml'><pkg:xmlData><w:document xmlns:w='http://schemas.openxmlformats.org/wordprocessingml/2006/main' ><w:body><w:p><w:pPr><w:spacing w:before='360' w:after='0' w:line='480' w:lineRule='auto'/><w:rPr><w:color w:val='70AD47' w:themeColor='accent6'/><w:sz w:val='28'/></w:rPr></w:pPr><w:r><w:rPr><w:color w:val='70AD47' w:themeColor='accent6'/><w:sz w:val='28'/></w:rPr><w:t>This text has formatting directly applied to achieve its font size, color, line spacing, and paragraph spacing.</w:t></w:r></w:p></w:body></w:document></pkg:xmlData></pkg:part></pkg:package>",
      Word.InsertLocation.start
    );

    // Synchronize the document state by executing the queued commands, and return a promise to indicate task completion.
    await context.sync();

    console.log("Added OOXML to the beginning of the document body.");
  });

  // Read "Understand when and how to use Office Open XML in your Word add-in" for guidance on working with OOXML.
  // https://learn.microsoft.com/office/dev/add-ins/word/create-better-add-ins-for-word-with-office-open-xml

  // The Word-Add-in-DocumentAssembly sample shows how you can use this API to assemble a document.
  // https://github.com/OfficeDev/Word-Add-in-DocumentAssembly
}

async function insertText() {
  // Inserts text at the beginning of this document.
  // Run a batch operation against the Word object model.
  await Word.run(async (context) => {
    // Create a proxy object for the document body.
    const body: Word.Body = context.document.body;

    // Queue a command to insert text at the beginning of the current document.
    body.insertText('This is text inserted with body.insertText()', Word.InsertLocation.start);

    // Synchronize the document state by executing the queued commands, and return a promise to indicate task completion.
    await context.sync();

    console.log("Text added to the beginning of the document body.");
  });
}

async function select() {
  // Selects the entire body.
  // Run a batch operation against the Word object model.
  await Word.run(async (context) => {
    // Create a proxy object for the document body.
    const body: Word.Body = context.document.body;

    // Queue a command to select the document body.
    // The Word UI will move to the selected document body.
    body.select();

    console.log("Selected the document body.");
  });
}

async function clear() {
  // Clears out the content from the document body.
  // Run a batch operation against the Word object model.
  await Word.run(async (context) => {
    // Create a proxy object for the document body.
    const body: Word.Body = context.document.body;

    // Queue a command to clear the contents of the body.
    body.clear();

    console.log("Cleared the body contents.");
  });

  // The Silly stories add-in sample shows how the clear method can be used to clear the contents of a document.
  // https://aka.ms/sillystorywordaddin
}

async function insertParagraph() {
  // Inserts a paragraph at the end of this document.
  // Run a batch operation against the Word object model.
  await Word.run(async (context) => {
    // Create a proxy object for the document body.
    const body: Word.Body = context.document.body;

    // Queue a command to insert a paragraph at the end of the current document.
    body.insertParagraph("Content of a new paragraph", Word.InsertLocation.end);

    // Synchronize the document state by executing the queued commands, and return a promise to indicate task completion.
    await context.sync();

    console.log("Paragraph added at the end of the document body.");
  });
  
  // The Word-Add-in-DocumentAssembly sample shows how you can use the insertParagraph method to assemble a document.
  // https://github.com/OfficeDev/Word-Add-in-DocumentAssembly
}

async function setup() {
  await Word.run(async (context) => {
    const body: Word.Body = context.document.body;
    body.clear();
    body.insertParagraph(
      "Do you want to create a solution that extends the functionality of Word? You can use the Office Add-ins platform to extend Word clients running on the web, on a Windows desktop, or on a Mac.",
      "Start"
    );
    body.paragraphs
      .getLast()
      .insertText(
        "Use add-in commands to extend the Word UI and launch task panes that run JavaScript that interacts with the content in a Word document. Any code that you can run in a browser can run in a Word add-in. Add-ins that interact with content in a Word document create requests to act on Word objects and synchronize object state.",
        "Replace"
      );
  });
}

// Default helper for invoking an action and handling errors.
async function tryCatch(callback) {
  try {
    await callback();
  } catch (error) {
    // Note: In a production add-in, you'd want to notify the user through your add-in's UI.
    console.error(error);
  }
}
```

