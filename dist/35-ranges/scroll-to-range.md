# Scroll to a range

Scrolls to a range with and without selection.

This sample demonstrates how to scroll to a range.

```typescript
async function scroll() {
  await Word.run(async (context) => {
    // If select is called with no parameters, it selects the object.
    context.document.body.paragraphs.getLast().select();

    await context.sync();
  });
}

async function scrollEnd() {
  await Word.run(async (context) => {
    // Select can be at the start or end of a range; this by definition moves the insertion point without selecting the range.
    context.document.body.paragraphs.getLast().select(Word.SelectionMode.end);

    await context.sync();
  });
}

async function setup() {
  await Word.run(async (context) => {
    const body: Word.Body = context.document.body;
    body.clear();
    const firstSentence: Word.Paragraph = body.insertParagraph(
      "Do you want to create a solution that extends the functionality of Word? You can use the Office Add-ins platform to extend Word clients running on the web, on a Windows desktop, or on a Mac.",
      "Start"
    );
    firstSentence.insertBreak(Word.BreakType.page, "After");
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

