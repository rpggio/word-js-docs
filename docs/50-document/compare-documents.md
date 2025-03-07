# Compare documents

Compares two documents (the current one and a specified external one).

This sample shows how to compare two documents: the current one and a specified external one.

```typescript
async function run() {
  // Compares the current document with a specified external document.
  await Word.run(async (context) => {
    // Absolute path of an online or local document.
    const filePath = $("#filePath")
      .val()
      .toString();
    // Options that configure the compare operation.
    const options: Word.DocumentCompareOptions = {
      compareTarget: Word.CompareTarget.compareTargetCurrent,
      detectFormatChanges: false
      // Other options you choose...
      };
    context.document.compare(filePath, options);

    await context.sync();

    console.log("Differences shown in the current document.");
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

