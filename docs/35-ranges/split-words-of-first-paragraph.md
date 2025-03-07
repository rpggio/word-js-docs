# Split a paragraph into ranges

Splits a paragraph into word ranges and then traverses all the ranges to format each word, producing a "karaoke" effect.

This sample demonstrates splitting and traversing ranges.

```typescript
async function highlightWords() {
  await Word.run(async (context) => {
    const paragraph: Word.Paragraph = context.document.body.paragraphs.getFirst();
    const words = paragraph.split([" "], true /* trimDelimiters*/, true /* trimSpaces */);
    words.load("text");

    await context.sync();

    for (let i = 0; i < words.items.length; i++) {
      if (i >= 1) {
        words.items[i - 1].font.highlightColor = "#FFFFFF";
      }
      words.items[i].font.highlightColor = "#FFFF00";

      await context.sync();
      await pause(200);
    }
  });
}

function pause(milliseconds) {
  return new Promise((resolve) => setTimeout(resolve, milliseconds));
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

