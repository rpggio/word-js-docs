# On adding content controls

Registers, triggers, and deregisters onAdded event that tracks the addition of content controls.

This sample demonstrates how to use the onAdded event with content controls.

```typescript
let eventContext;

async function registerEventHandler() {
  // Registers the onAdded event handler on the document.
  await Word.run(async (context) => {
    eventContext = context.document.onContentControlAdded.add(contentControlAdded);
    await context.sync();

    console.log("Added event handler for when content controls are added.");
  });
}

async function contentControlAdded(event: Word.ContentControlAddedEventArgs) {
  await Word.run(async (context) => {
    console.log(`${event.eventType} event detected. IDs of content controls that were added:`, event.ids);
  });
}

async function insertContentControls() {
  // Traverses each paragraph of the document and wraps a content control on each.
  await Word.run(async (context) => {
    const paragraphs: Word.ParagraphCollection = context.document.body.paragraphs;
    paragraphs.load("$none"); // Don't need any properties; just wrap each paragraph with a content control.

    await context.sync();

    for (let i = 0; i < paragraphs.items.length; i++) {
      let contentControl = paragraphs.items[i].insertContentControl();
      contentControl.tag = "forTesting";
    }

    console.log("Content controls inserted: " + paragraphs.items.length);

    await context.sync();
  });
}

async function deregisterEventHandler() {
    await Word.run(eventContext.context, async (context) => {
      eventContext.remove();
      await context.sync();
    });

  eventContext = null;
  console.log("Removed event handler that was tracking when content controls are added.");
}

async function setup() {
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

