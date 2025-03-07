# Manage footnotes

This sample shows how to perform basic footnote operations, including insert, get, and delete.

This sample shows basic operations using footnotes.

```typescript
async function insertFootnote() {
  // Sets a footnote on the selected content.
  await Word.run(async (context) => {
    const text = $("#input-footnote")
      .val()
      .toString();
    const footnote: Word.NoteItem = context.document.getSelection().insertFootnote(text);
    await context.sync();

    console.log("Inserted footnote.");
  });
}
async function getReference() {
  // Selects the footnote's reference mark in the document body.
  await Word.run(async (context) => {
    const footnotes: Word.NoteItemCollection = context.document.body.footnotes;
    footnotes.load("items/reference");
    await context.sync();

    const referenceNumber = $("#input-reference").val();
    const mark = (referenceNumber as number) - 1;
    const item: Word.NoteItem = footnotes.items[mark];
    const reference: Word.Range = item.reference;
    reference.select();
    await context.sync();

    console.log(`Reference ${referenceNumber} is selected.`);
  });
}
async function getFootnoteType() {
  // Gets the referenced note's item type and body type, which are both "Footnote".
  await Word.run(async (context) => {
    const footnotes: Word.NoteItemCollection = context.document.body.footnotes;
    footnotes.load("items");
    await context.sync();

    const referenceNumber = $("#input-reference").val();
    const mark = (referenceNumber as number) - 1;
    const item: Word.NoteItem = footnotes.items[mark];
    console.log(`Note type of footnote ${referenceNumber}: ${item.type}`);

    item.body.load("type");
    await context.sync();

    console.log(`Body type of note: ${item.body.type}`);
  });
}
async function getFootnoteBody() {
  // Gets the text of the referenced footnote.
  await Word.run(async (context) => {
    const footnotes: Word.NoteItemCollection = context.document.body.footnotes;
    footnotes.load("items/body");
    await context.sync();

    const referenceNumber = $("#input-reference").val();
    const mark = (referenceNumber as number) - 1;
    const footnoteBody: Word.Range = footnotes.items[mark].body.getRange();
    footnoteBody.load("text");
    await context.sync();

    console.log(`Text of footnote ${referenceNumber}: ${footnoteBody.text}`);
  });
}
async function getNextFootnote() {
  // Selects the next footnote in the document body.
  await Word.run(async (context) => {
    const footnotes: Word.NoteItemCollection = context.document.body.footnotes;
    footnotes.load("items/reference");
    await context.sync();

    const referenceNumber = $("#input-reference").val();
    const mark = (referenceNumber as number) - 1;
    const reference: Word.Range = footnotes.items[mark].getNext().reference;
    reference.select();
    console.log("Selected is the next footnote: " + (mark + 2));
  });
}
async function deleteFootnote() {
  // Deletes this referenced footnote.
  await Word.run(async (context) => {
    const footnotes: Word.NoteItemCollection = context.document.body.footnotes;
    footnotes.load("items");
    await context.sync();

    const referenceNumber = $("#input-reference").val();
    const mark = (referenceNumber as number) - 1;
    footnotes.items[mark].delete();
    await context.sync();

    console.log("Footnote deleted.");
  });
}
async function getFirstFootnote() {
  // Gets the first footnote in the document body and select its reference mark.
  await Word.run(async (context) => {
    const reference: Word.Range = context.document.body.footnotes.getFirst().reference;
    reference.select();
    console.log("The first footnote is selected.");
  });
}
async function getFootnotesFromBody() {
  // Gets the footnotes in the document body.
  await Word.run(async (context) => {
    const footnotes: Word.NoteItemCollection = context.document.body.footnotes;
    footnotes.load("length");
    await context.sync();

    console.log("Number of footnotes in the document body: " + footnotes.items.length);
  });
}
async function getFootnotesFromRange() {
  // Gets the footnotes in the selected document range.
  await Word.run(async (context) => {
    const footnotes: Word.NoteItemCollection = context.document.getSelection().footnotes;
    footnotes.load("length");
    await context.sync();

    console.log("Number of footnotes in the selected range: " + footnotes.items.length);
  });
}
async function setup() {
  // Set two paragraphs of sample text.
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

