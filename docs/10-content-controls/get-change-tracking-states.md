# Get change tracking states of content controls

Gets change tracking states of content controls.

This sample demonstrates how to insert and delete control controls then get their change tracking state.

```typescript
async function turnOnChangeTracking() {
  // Turns on change tracking.
  await Word.run(async (context) => {
    context.document.changeTrackingMode = Word.ChangeTrackingMode.trackAll;
    await context.sync();
    console.log("Turned on change tracking.");
  });
}

async function insertContentControlOnLastParagraph() {
  // Inserts a content control on the last paragraph.
  await Word.run(async (context) => {
    let paragraphs = context.document.body.paragraphs;
    paragraphs.load();
    paragraphs.getLast().insertContentControl("RichText");
    await context.sync();
    console.log("Inserted content control on last paragraph.");
  });
}

async function deleteFirstContentControl() {
  // Deletes the first content control found in the document body.
  await Word.run(async (context) => {
    context.document.body
      .getContentControls()
      .getFirst()
      .delete(false);
    await context.sync();
    console.log("Deleted the first content control.");
  });
}

async function turnOffChangeTracking() {
  // Turns off change tracking.
  await Word.run(async (context) => {
    context.document.changeTrackingMode = Word.ChangeTrackingMode.off;
    await context.sync();
    console.log("Turned off change tracking.");
  });
}

async function logChangeTrackingStates() {
  // Logs the current change tracking states of the content controls.
  await Word.run(async (context) => {
    let trackAddedArray: Word.ChangeTrackingState[] = [Word.ChangeTrackingState.added];
    let trackDeletedArray: Word.ChangeTrackingState[] = [Word.ChangeTrackingState.deleted];
    let trackNormalArray: Word.ChangeTrackingState[] = [Word.ChangeTrackingState.normal];

    let addedContentControls = context.document.body.getContentControls().getByChangeTrackingStates(trackAddedArray);
    let deletedContentControls = context.document.body
      .getContentControls()
      .getByChangeTrackingStates(trackDeletedArray);
    let normalContentControls = context.document.body.getContentControls().getByChangeTrackingStates(trackNormalArray);

    addedContentControls.load();
    deletedContentControls.load();
    normalContentControls.load();
    await context.sync();

    console.log(`Number of content controls in Added state: ${addedContentControls.items.length}`);
    console.log(`Number of content controls in Deleted state: ${deletedContentControls.items.length}`);
    console.log(`Number of content controls in Normal state: ${normalContentControls.items.length}`);
  });
}

async function setup() {
  // Adds 4 paragraphs then inserts content controls on the first three paragraphs.
  await Word.run(async (context) => {
    const body: Word.Body = context.document.body;
    body.clear();
    body.insertParagraph("One more paragraph.", "Start");
    body.insertParagraph("Inserting another paragraph.", "Start");
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
    body.paragraphs.load();
    await context.sync();
    body.paragraphs.items[0].insertContentControl("PlainText");
    body.paragraphs.items[1].insertContentControl("PlainText");
    body.paragraphs.items[2].insertContentControl("RichText");
    console.log("Inserted content controls on the first three paragraphs.");
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

