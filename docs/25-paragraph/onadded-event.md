# On adding paragraphs

Registers, triggers, and deregisters the onParagraphAdded event that tracks the addition of paragraphs.

This sample demonstrates how to use the onAdded event with paragraphs.

```typescript
let eventContext;

async function registerEventHandler() {
  // Registers the onParagraphAdded event handler on the document.
  await Word.run(async (context) => {
    eventContext = context.document.onParagraphAdded.add(paragraphAdded);
    await context.sync();

    console.log("Added event handler for when paragraphs are added.");
  });
}

async function paragraphAdded(event: Word.ParagraphAddedEventArgs) {
  await Word.run(async (context) => {
    console.log(`${event.type} event detected. IDs of paragraphs that were added:`, event.uniqueLocalIds);
  });
}

async function getParagraphById() {
  await Word.run(async (context) => {
    const paragraphId = $("#paragraph-id").val() as string;
    const paragraph: Word.Paragraph = context.document.getParagraphByUniqueLocalId(paragraphId);
    paragraph.load();
    await paragraph.context.sync();

    console.log(paragraph);
  });
}

async function insertParagraphs() {
  // Inserts two paragraphs within the document body.
  await Word.run(async (context) => {
    const paragraphCount = 2;
    for (let i = 0; i < paragraphCount; i++) {
      context.document.body.insertParagraph(`Paragraph Test ${i + 1}`, "End");
    }

    console.log("Paragraphs inserted: " + paragraphCount);
    await context.sync();
  });
}

async function deregisterEventHandler() {
  await Word.run(eventContext.context, async (context) => {
    eventContext.remove();
    await context.sync();
  });

  eventContext = null;
  console.log("Removed event handler that was tracking when paragraphs are added.");
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

