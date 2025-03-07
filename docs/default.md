# Blank snippet

Creates a new snippet from a blank template.

**Button:** Run

```typescript
async function run() {
    await Word.run(async (context) => {
        const body = context.document.body;

        console.log("Your code goes here");

        await context.sync();
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

