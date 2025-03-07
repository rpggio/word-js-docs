# Compare range locations

This sample shows how to compare the locations of two ranges.

This sample demonstrates how to compare locations of ranges.

```typescript
async function compareLocations() {
  // Compares the location of one paragraph in relation to another paragraph.
  await Word.run(async (context) => {
    const paragraphs: Word.ParagraphCollection = context.document.body.paragraphs;
    paragraphs.load("items");

    await context.sync();

    const firstParagraphAsRange: Word.Range = paragraphs.items[0].getRange();
    const secondParagraphAsRange: Word.Range = paragraphs.items[1].getRange();

    const comparedLocation = firstParagraphAsRange.compareLocationWith(secondParagraphAsRange);

    await context.sync();

    const locationValue: Word.LocationRelation = comparedLocation.value;
    console.log(`Location of the first paragraph in relation to the second paragraph: ${locationValue}`);
  });
}

async function compareWithSelection() {
  // Compares the location of the second paragraph in relation to the cursor's location.
  await Word.run(async (context) => {
    const paragraphs: Word.ParagraphCollection = context.document.body.paragraphs;
    paragraphs.load("items");

    await context.sync();

    const secondParagraphAsRange: Word.Range = paragraphs.items[1].getRange();
    const selection: Word.Range = context.document.getSelection();

    const comparedLocation = secondParagraphAsRange.compareLocationWith(selection);

    await context.sync();

    const locationValue: Word.LocationRelation = comparedLocation.value;
    console.log(`Location of the second paragraph in relation to the cursor's location: ${locationValue}`);
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

