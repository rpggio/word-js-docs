# Insert an external document

Inserts the content (with or without settings) of an external document into the current document. Settings include formatting, change-tracking mode, custom properties, and XML parts.

This sample shows how to insert the body text from an external document into the current document.

```typescript
let externalDocument;

async function insertDocument() {
  // Updates the text of the current document with the text from another document passed in as a Base64-encoded string.
  await Word.run(async (context) => {
    // Use the Base64-encoded string representation of the selected .docx file.
    const externalDoc: Word.DocumentCreated = context.application.createDocument(externalDocument);
    await context.sync();

    if (!Office.context.requirements.isSetSupported("WordApiHiddenDocument", "1.3")) {
      console.warn("The WordApiHiddenDocument 1.3 requirement set isn't supported on this client so can't proceed. Try this action on a platform that supports this requirement set.");
      return;
    }

    const externalDocBody: Word.Body = externalDoc.body;
    externalDocBody.load("text");
    await context.sync();

    // Insert the external document's text at the beginning of the current document's body.
    const externalDocBodyText = externalDocBody.text;
    const currentDocBody: Word.Body = context.document.body;
    currentDocBody.insertText(externalDocBodyText, Word.InsertLocation.start);
    await context.sync();
  });
}

async function insertDocumentWithSettings() {
  // Inserts content (applying selected settings) from another document passed in as a Base64-encoded string.
  await Word.run(async (context) => {
    // Use the Base64-encoded string representation of the selected .docx file.
    context.document.insertFileFromBase64(externalDocument, "Replace", {
      importTheme: true,
      importStyles: true,
      importParagraphSpacing: true,
      importPageColor: true,
      importChangeTrackingMode: true,
      importCustomProperties: true,
      importCustomXmlParts: true,
      importDifferentOddEvenPages: true
    });
    await context.sync();
  });
}

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

