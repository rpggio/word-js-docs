# Correlated objects pattern

Shows the performance benefits of avoiding `context.sync` calls in a loop.

This sample demonstrates the performance optimization gained from the correlated objects pattern. For more information, see Avoid using the context.sync method in loops.

```typescript
const jobMapping = [
  { job: "{Coordinator}", person: "Sally" },
  { job: "{Deputy}", person: "Bob" },
  { job: "{Manager}", person: "Kim" }
];
async function replacePlaceholders() {
  Word.run(async (context) => {
    const startTime = Date.now();
    let count = 0;

    // Find the locations of all the placeholder strings.
    const allSearchResults = [];
    for (let i = 0; i < jobMapping.length; i++) {
      let options = Word.SearchOptions.newObject(context);
      options.matchWildcards = false;
      let searchResults = context.document.body.search(jobMapping[i].job, options);
      searchResults.load('items');
      let correlatedSearchResult = {
        rangesMatchingJob: searchResults,
        personAssignedToJob: jobMapping[i].person
      }
      allSearchResults.push(correlatedSearchResult);
    }

    // Sync to load those locations in the add-in.
    await context.sync()

    // Replace the placeholder text at the known locations.
    for (let i = 0; i < allSearchResults.length; i++) {
      let correlatedObject = allSearchResults[i];

      for (let j = 0; j < correlatedObject.rangesMatchingJob.items.length; j++) {
        let targetRange = correlatedObject.rangesMatchingJob.items[j];
        let name = correlatedObject.personAssignedToJob;
        targetRange.insertText(name, Word.InsertLocation.replace);
        count++;
      }
    }

    await context.sync();
    console.log(`Replacing ${count} placeholders with the correlated objects pattern took ${Date.now() - startTime} milliseconds.`);
    console.log()
  });
}
async function replacePlaceholdersSlow() {
  Word.run(async (context) => {
    const startTime = Date.now();
    let count = 0;

    // The context.sync calls in the loops will degrade performance.
    for (let i = 0; i < jobMapping.length; i++) {
      let options = Word.SearchOptions.newObject(context);
      options.matchWildcards = false;
      let searchResults = context.document.body.search(jobMapping[i].job, options);
      searchResults.load('items');

      await context.sync();

      for (let j = 0; j < searchResults.items.length; j++) {
        searchResults.items[j].insertText(jobMapping[i].person, Word.InsertLocation.replace);
        count++;
        await context.sync();
      }
    }
    console.log(`Replacing ${count} placeholders with in-loop sync statements took ${Date.now() - startTime} milliseconds.`);
  });
}
async function setup(timesToAddText: number = 1) {
  await Word.run(async (context) => {
    console.log("Setup beginning...");
    const body: Word.Body = context.document.body;
    body.clear();
    while (timesToAddText > 0) {
      body.insertParagraph(
        "This defines the roles of {Coordinator}, {Deputy}, {Manager}.",
        Word.InsertLocation.end
      );
      body.insertParagraph(
        "{Coordinator}: Oversees daily operations and ensures projects run smoothly by coordinating between different teams and resources.",
        Word.InsertLocation.end
      );
      body.insertParagraph(
        "{Deputy}: Assists and supports senior management, often stepping in to make decisions or manage tasks in {Manager}'s absence.",
        Word.InsertLocation.end
      );
      body.insertParagraph(
        "{Manager}: Leads the team, setting goals, planning strategies, and making decisions to achieve organizational objectives.",
        Word.InsertLocation.end
      );
      timesToAddText--;
    }
    await context.sync();
    console.log("Setup complete.");
  });
}
async function addLotsOfText() {
  // Add the setup text 100 times.
  setup(100);
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

