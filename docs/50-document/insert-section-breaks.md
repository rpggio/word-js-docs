# Add a section

Shows how to insert section breaks in the document.

This sample shows how to insert sections in the document.

```typescript
async function addNext() {
  // Inserts a section break on the next page.
  await Word.run(async (context) => {
    const body: Word.Body = context.document.body;
    body.insertBreak(Word.BreakType.sectionNext, Word.InsertLocation.end);

    await context.sync();

    console.log("Inserted section break on next page.");
  });
}

async function addEven() {
  // Inserts a section break on the next even page.
  await Word.run(async (context) => {
    const body: Word.Body = context.document.body;
    body.insertBreak(Word.BreakType.sectionEven, Word.InsertLocation.end);

    await context.sync();

    console.log("Inserted section break on next even page.");
  });
}

async function addOdd() {
  // Inserts a section break on the next odd page.
  await Word.run(async (context) => {
    const body: Word.Body = context.document.body;
    body.insertBreak(Word.BreakType.sectionOdd, Word.InsertLocation.end);

    await context.sync();

    console.log("Inserted section break on next odd page.");
  });
}

async function addContinuous() {
  // Inserts a section without an associated page break.
  await Word.run(async (context) => {
    const body: Word.Body = context.document.body;
    body.insertBreak(Word.BreakType.sectionContinuous, Word.InsertLocation.end);

    await context.sync();

    console.log("Inserted section without an associated page break.");
  });
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

