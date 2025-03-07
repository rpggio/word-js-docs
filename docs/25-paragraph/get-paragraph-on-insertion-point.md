# Get paragraph from insertion point

Gets the full paragraph containing the insertion point.

This sample demonstrates how to get the paragraph and paragraph sentences associated with the current insertion point.

```typescript
async function getParagraph() {
  await Word.run(async (context) => {
    // The collection of paragraphs of the current selection returns the full paragraphs contained in it.
    const paragraph: Word.Paragraph = context.document.getSelection().paragraphs.getFirst();
    paragraph.load("text");

    await context.sync();
    console.log(paragraph.text);
  });
}

async function getSentences() {
  await Word.run(async (context) => {
    // Get the complete sentence (as range) associated with the insertion point.
    const sentences: Word.RangeCollection = context.document
      .getSelection()
      .getTextRanges(["."] /* Using the "." as delimiter */, false /*means without trimming spaces*/);
    sentences.load("$none");
    await context.sync();

    // Expand the range to the end of the paragraph to get all the complete sentences.
    const sentencesToTheEndOfParagraph: Word.RangeCollection = sentences.items[0]
      .getRange()
      .expandTo(
        context.document
          .getSelection()
          .paragraphs.getFirst()
          .getRange(Word.RangeLocation.end)
      )
      .getTextRanges(["."], false /* Don't trim spaces*/);
    sentencesToTheEndOfParagraph.load("text");
    await context.sync();

    for (let i = 0; i < sentencesToTheEndOfParagraph.items.length; i++) {
      console.log(sentencesToTheEndOfParagraph.items[i].text);
    }
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

