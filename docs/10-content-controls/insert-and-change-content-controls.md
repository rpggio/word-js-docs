# Content control basics

Inserts, updates, and retrieves content controls.

This sample demonstrates how to insert and change content control properties.

```typescript
async function insertContentControls() {
  // Traverses each paragraph of the document and wraps a content control on each with either a even or odd tags.
  await Word.run(async (context) => {
    let paragraphs = context.document.body.paragraphs;
    paragraphs.load("$none"); // Don't need any properties; just wrap each paragraph with a content control.

    await context.sync();

    for (let i = 0; i < paragraphs.items.length; i++) {
      let contentControl = paragraphs.items[i].insertContentControl();
      // For even, tag "even".
      if (i % 2 === 0) {
        contentControl.tag = "even";
      } else {
        contentControl.tag = "odd";
      }
    }
    console.log("Content controls inserted: " + paragraphs.items.length);

    await context.sync();
  });
}

async function modifyContentControls() {
  // Adds title and colors to odd and even content controls and changes their appearance.
  await Word.run(async (context) => {
    // Get the complete sentence (as range) associated with the insertion point.
    let evenContentControls = context.document.contentControls.getByTag("even");
    let oddContentControls = context.document.contentControls.getByTag("odd");
    evenContentControls.load("length");
    oddContentControls.load("length");

    await context.sync();

    for (let i = 0; i < evenContentControls.items.length; i++) {
      // Change a few properties and append a paragraph.
      evenContentControls.items[i].set({
        color: "red",
        title: "Odd ContentControl #" + (i + 1),
        appearance: Word.ContentControlAppearance.tags
      });
      evenContentControls.items[i].insertParagraph("This is an odd content control", "End");
    }

    for (let j = 0; j < oddContentControls.items.length; j++) {
      // Change a few properties and append a paragraph.
      oddContentControls.items[j].set({
        color: "green",
        title: "Even ContentControl #" + (j + 1),
        appearance: "Tags"
      });
      oddContentControls.items[j].insertHtml("This is an <b>even</b> content control", "End");
    }

    await context.sync();
  });
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

