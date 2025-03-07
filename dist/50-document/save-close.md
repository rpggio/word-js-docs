# Manage document save and close

Shows how to manage saving and closing document.

This sample shows how to use the options for saving and closing the current document.

```typescript
async function saveNoPrompt() {
  // Saves the document with the provided file name
  // if it hasn't been saved before.
  await Word.run(async (context) => {
    const text = $("#fileName-text")
      .val()
      .toString();
    context.document.save(Word.SaveBehavior.save, text);
    await context.sync();
  });
}

async function saveAfterPrompt() {
  // If the document hasn't been saved before, prompts
  // user with options for if or how they want to save.
  await Word.run(async (context) => {
    context.document.save(Word.SaveBehavior.prompt);
    await context.sync();
  });
}

async function closeAfterSave() {
  // Closes the document after saving.
  await Word.run(async (context) => {
    context.document.close(Word.CloseBehavior.save);
  });
}

async function closeWithoutSave() {
  // Closes the document without saving any changes.
  await Word.run(async (context) => {
    context.document.close(Word.CloseBehavior.skipSave);
    await context.sync();
  });
}

async function save() {
  // Saves the document with default behavior
  // for current state of the document.
  await Word.run(async (context) => {
    context.document.save();
    await context.sync();
  });
}

async function close() {
  // Closes the document with default behavior
  // for current state of the document.
  await Word.run(async (context) => {
    context.document.close();
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

