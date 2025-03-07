# Manage tracked changes

This samples shows how to manage tracked changes, including accepting and rejecting changes.

This sample shows how to manage tracked changes.

```typescript
async function getAllTrackedChanges() {
  // Gets all tracked changes.
  await Word.run(async (context) => {
    const body: Word.Body = context.document.body;
    const trackedChanges: Word.TrackedChangeCollection = body.getTrackedChanges();
    trackedChanges.load();
    await context.sync();

    console.log(trackedChanges);
  });
}

async function getFirstTrackedChangeRange() {
  // Gets the range of the first tracked change.
  await Word.run(async (context) => {
    const body: Word.Body = context.document.body;
    const trackedChanges: Word.TrackedChangeCollection = body.getTrackedChanges();
    const trackedChange: Word.TrackedChange = trackedChanges.getFirst();
    await context.sync();

    const range: Word.Range = trackedChange.getRange();
    range.load();
    await context.sync();

    console.log("range.text: " + range.text);
  });
}

async function getNextTrackedChange() {
  // Gets the next (second) tracked change.
  await Word.run(async (context) => {
    const body: Word.Body = context.document.body;
    const trackedChanges: Word.TrackedChangeCollection = body.getTrackedChanges();
    await context.sync();

    const trackedChange: Word.TrackedChange = trackedChanges.getFirst();
    await context.sync();

    const nextTrackedChange: Word.TrackedChange = trackedChange.getNext();
    await context.sync();

    nextTrackedChange.load(["author", "date", "text", "type"]);
    await context.sync();

    console.log(nextTrackedChange);
  });
}

async function acceptFirstTrackedChange() {
  // Accepts the first tracked change.
  await Word.run(async (context) => {
    const body: Word.Body = context.document.body;
    const trackedChanges: Word.TrackedChangeCollection = body.getTrackedChanges();
    const trackedChange: Word.TrackedChange = trackedChanges.getFirst();
    trackedChange.load();
    await context.sync();

    console.log("First tracked change:", trackedChange);
    trackedChange.accept();
    console.log("Accepted the first tracked change.");
  });
}

async function rejectFirstTrackedChange() {
  // Rejects the first tracked change.
  await Word.run(async (context) => {
    const body: Word.Body = context.document.body;
    const trackedChanges: Word.TrackedChangeCollection = body.getTrackedChanges();
    const trackedChange: Word.TrackedChange = trackedChanges.getFirst();
    trackedChange.load();
    await context.sync();

    console.log("First tracked change:", trackedChange);
    trackedChange.reject();
    console.log("Rejected the first tracked change.");
  });
}

async function acceptAllTrackedChanges() {
  // Accepts all tracked changes.
  await Word.run(async (context) => {
    const body: Word.Body = context.document.body;
    const trackedChanges: Word.TrackedChangeCollection = body.getTrackedChanges();
    trackedChanges.acceptAll();
    console.log("Accepted all tracked changes.");
  });
}

async function rejectAllTrackedChanges() {
  // Rejects all tracked changes.
  await Word.run(async (context) => {
    const body: Word.Body = context.document.body;
    const trackedChanges: Word.TrackedChangeCollection = body.getTrackedChanges();
    trackedChanges.rejectAll();
    console.log("Rejected all tracked changes.");
  });
}

async function setup() {
  // Updates the text and sets the font color to red.
  await Word.run(async (context) => {
    context.document.changeTrackingMode = Word.ChangeTrackingMode.off;

    context.document.body.insertText("AAA BBB CCC DDD EEE FFF", "Replace");

    context.document.changeTrackingMode = Word.ChangeTrackingMode.trackAll;
    context.document.body
      .search("BBB")
      .getFirst()
      .insertText("WWW", "Replace");
    context.document.body
      .search("DDD ")
      .getFirst()
      .delete();
    context.document.body
      .search("FFF")
      .getFirst()
      .insertText("XXX ", "Start");
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

