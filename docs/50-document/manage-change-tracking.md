# Track changes

This sample shows how to get and set the change tracking mode and get the before and after of reviewed text.

This sample shows basic operations of the Track Changes feature.

```typescript
async function getChangeTrackingMode() {
  // Gets the current change tracking mode.
  await Word.run(async (context) => {
    const document: Word.Document = context.document;
    document.load("changeTrackingMode");
    await context.sync();

    if (document.changeTrackingMode === Word.ChangeTrackingMode.trackMineOnly) {
      console.log("Only my changes are being tracked.");
    } else if (document.changeTrackingMode === Word.ChangeTrackingMode.trackAll) {
      console.log("Everyone's changes are being tracked.");
    } else {
      console.log("No changes are being tracked.");
    }
  });
}

async function setChangeTrackingMode() {
  // Sets the change tracking mode.
  await Word.run(async (context) => {
    const mode = $("input[name='mode']:checked").val();
    if (mode === "Track only my changes") {
      context.document.changeTrackingMode = Word.ChangeTrackingMode.trackMineOnly;
    } else if (mode === "Track everyone's changes") {
      context.document.changeTrackingMode = Word.ChangeTrackingMode.trackAll;
    } else {
      context.document.changeTrackingMode = Word.ChangeTrackingMode.off;
    }

    await context.sync();

    getChangeTrackingMode();
  });
}

async function getReviewedText() {
  // Gets the reviewed text.
  await Word.run(async (context) => {
    const range: Word.Range = context.document.getSelection();
    const before = range.getReviewedText(Word.ChangeTrackingVersion.original);
    const after = range.getReviewedText(Word.ChangeTrackingVersion.current);

    await context.sync();

    console.log("Reviewed text (before):", before.value, "Reviewed text (after):", after.value);
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

