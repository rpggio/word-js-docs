# Basic API call (TypeScript)

Performs a basic Word API call using TypeScript.

This sample executes a code snippet that prints the selected text to the console. Make sure to enter and select text before clicking "Print selection".

**Button:** Print selection

```typescript
async function run() {
    // Gets the current selection and changes the font color to red.
    await Word.run(async (context) => {
        const range: Word.Range = context.document.getSelection();
        range.font.color = "red";
        range.load("text");

        await context.sync();

        console.log(`The selected text was "${range.text}".`);
    });
}

// Default helper for invoking an action and handling errors.
async function tryCatch(callback) {
    try {
        await callback();
    }
    catch (error) {
        // Note: In a production add-in, you'd want to notify the user through your add-in's UI.
        console.error(error);
    }
}
```

