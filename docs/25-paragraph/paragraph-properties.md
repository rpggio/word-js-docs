# Paragraph properties

Sets indentation, space between paragraphs, and other paragraph properties.

This sample demonstrates paragraph property usage.

```typescript
async function indent() {
  await Word.run(async (context) => {
    // Indent the first paragraph.
    context.document.body.paragraphs.getFirst().leftIndent = 75; //units = points

    return context.sync();
  });
}

async function spacing() {
  await Word.run(async (context) => {
    // Adjust line spacing.
    context.document.body.paragraphs.getFirst().lineSpacing = 20;

    await context.sync();
  });
}

async function spaceAfter() {
  await Word.run(async (context) => {
    // Set the space (in points) after the first paragraph.
    context.document.body.paragraphs.getFirst().spaceAfter = 20;

    await context.sync();
  });
}

async function spaceAfterInLines() {
  await Word.run(async (context) => {
    // Set the space (in line units) after the first paragraph.
    context.document.body.paragraphs.getFirst().lineUnitAfter = 1;

    await context.sync();
  });
}

async function spaceBeforeInLines() {
  await Word.run(async (context) => {
    // Set the space (in line units) before the first paragraph.
    context.document.body.paragraphs.getFirst().lineUnitBefore = 1;

    await context.sync();
  });
}

async function align() {
  await Word.run(async (context) => {
    // Center last paragraph alignment.
    context.document.body.paragraphs.getLast().alignment = "Centered";

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
    body.paragraphs.getFirst().alignment = "Left";
    body.paragraphs.getLast().alignment = Word.Alignment.left;
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

