# Get word count

Counts how many times a word or term appears in the document.

This sample demonstrates how to get the count for words and terms in the document body.

```typescript
async function run() {
  // Counts how many times each term appears in the document.
  await Word.run(async (context) => {
    const paragraphs: Word.ParagraphCollection = context.document.body.paragraphs;
    paragraphs.load("text");
    await context.sync();

    // Split up the document text using existing spaces as the delimiter.
    let text = [];
    paragraphs.items.forEach((item) => {
      let paragraph = item.text.trim();
      if (paragraph) {
        paragraph.split(" ").forEach((term) => {
          let currentTerm = term.trim();
          if (currentTerm) {
            text.push(currentTerm);
          }
        });
      }
    });

    // Determine the list of unique terms.
    let makeTextDistinct = new Set(text);
    let distinctText = Array.from(makeTextDistinct);
    let allSearchResults = [];

    for (let i = 0; i < distinctText.length; i++) {
      let results = context.document.body.search(distinctText[i], { matchCase: true, matchWholeWord: true });
      results.load("text");

      // Map each search term with its results.
      let correlatedResults = {
        searchTerm: distinctText[i],
        hits: results
      };

      allSearchResults.push(correlatedResults);
    }

    await context.sync();

    // Display the count for each search term.
    allSearchResults.forEach((result) => {
      let length = result.hits.items.length;

      console.log("Search term: " + result.searchTerm + " => Count: " + length);
    });
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
    body.insertParagraph(
      "Use add-in commands to extend the Word UI and launch task panes that run JavaScript that interacts with the content in a Word document. Any code that you can run in a browser can run in a Word add-in. Add-ins that interact with content in a Word document create requests to act on Word objects and synchronize object state.",
      "End"
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

