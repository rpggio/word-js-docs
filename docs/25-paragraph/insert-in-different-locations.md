# Insert content at different locations

Inserts content at different document locations.

This sample demonstrates a variety of insert locations available in the API.

```typescript
async function before() {
  await Word.run(async (context) => {
    // Let's insert before the first paragraph.
    const range: Word.Paragraph = context.document.body.paragraphs.getFirst().insertParagraph("This is Before", "Before");
    range.font.highlightColor = "yellow";

    await context.sync();
  });
}

async function start() {
  await Word.run(async (context) => {
    // This button assumes before() ran.
    // Get the next paragraph and insert text at the beginning. Note that there are invalid locations depending on the object. For instance, insertParagraph and "before" on a paragraph object is not a valid combination.
    const range: Word.Range = context.document.body.paragraphs
      .getFirst()
      .getNext()
      .insertText("This is Start", "Start");
    range.font.highlightColor = "blue";
    range.font.color = "white";

    await context.sync();
  });
}

async function end() {
  await Word.run(async (context) => {
    // Insert text at the end of a paragraph.
    const range: Word.Range = context.document.body.paragraphs
      .getFirst()
      .getNext()
      .insertText(" This is End", "End");
    range.font.highlightColor = "green";
    range.font.color = "white";

    await context.sync();
  });
}

async function after() {
  await Word.run(async (context) => {
    // Insert a paragraph after an existing one.
    const range: Word.Paragraph = context.document.body.paragraphs
      .getFirst()
      .getNext()
      .insertParagraph("This is After", "After");
    range.font.highlightColor = "red";
    range.font.color = "white";

    await context.sync();
  });
}

async function replace() {
  await Word.run(async (context) => {
    // Replace the last paragraph.
    const range: Word.Range = context.document.body.paragraphs.getLast().insertText("Just replaced the last paragraph!", "Replace");
    range.font.highlightColor = "black";
    range.font.color = "white";

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

async function tryCatch(callback) {
  try {
    await callback();
  } catch (error) {
    // Note: In a production add-in, you'd want to notify the user through your add-in's UI.
    console.error(error);
  }
}
```

