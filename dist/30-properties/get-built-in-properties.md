# Built-in document properties

Gets built-in document properties.

This sample demonstrates how to get the built-in properties of a Word document.

```typescript
async function getProperties() {
    await Word.run(async (context) => {
        const builtInProperties: Word.DocumentProperties = context.document.properties;
        builtInProperties.load("*"); // Let's get all!

        await context.sync();
        console.log(JSON.stringify(builtInProperties, null, 4));
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

