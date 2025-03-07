# Set multiple properties at once

Sets multiple properties at once with the API object set() method.

This sample shows how to format text using the object.set method.

```typescript
async function setMultiplePropertiesWithObject() {
  await Word.run(async (context) => {
    const paragraph: Word.Paragraph = context.document.body.paragraphs.getFirst();
    paragraph.set({
      leftIndent: 30,
      font: {
        bold: true,
        color: "red"
      }
    });

    await context.sync();
  });
}

async function copyPropertiesFromParagraph() {
  await Word.run(async (context) => {
    const firstParagraph: Word.Paragraph = context.document.body.paragraphs.getFirst();
    const secondParagraph: Word.Paragraph = firstParagraph.getNext();
    firstParagraph.load("text, font/color, font/bold, leftIndent");

    await context.sync();

    secondParagraph.set(firstParagraph);

    await context.sync();
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

