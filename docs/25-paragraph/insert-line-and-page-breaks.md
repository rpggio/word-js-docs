# Insert breaks

Inserts page and line breaks in a document.

This sample demonstrates how to insert page and line breaks.

```typescript
async function insertLineBreak() {
  Word.run(async (context) => {
    context.document.body.paragraphs.getFirst().insertBreak(Word.BreakType.line, "After");

    await context.sync();
    console.log("success");
  });
}

async function insertPageBreak() {
  await Word.run(async (context) => {
    context.document.body.paragraphs.getFirst().insertBreak(Word.BreakType.page, "After");

    await context.sync();
    console.log("success");
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

    console.log("success");
  });
}

async function tryCatch(callback) {
  try {
    await callback();
  } catch (error) {
    // Note: In a production add-in, you'd want to notify the user through your add-in's UI.
    console.error(error);
  }
}
```

