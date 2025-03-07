# Custom document properties

Adds and reads custom document properties of different types.

This sample demonstrates how to insert custom document properties of different data types and how to read them.

```typescript
async function insertNumericProperty() {
    await Word.run(async (context) => {
        context.document.properties.customProperties.add("Numeric Property", 1234);

        await context.sync();
        console.log("Property added");
    });
}

async function insertStringProperty() {
    await Word.run(async (context) => {
        context.document.properties.customProperties.add("String Property", "Hello World!");

        await context.sync();
        console.log("Property added");
    });
}

async function readCustomDocumentProperties() {
    await Word.run(async (context) => {
        const properties: Word.CustomPropertyCollection = context.document.properties.customProperties;
        properties.load("key,type,value");

        await context.sync();
        for (let i = 0; i < properties.items.length; i++)
            console.log("Property Name:" + properties.items[i].key + "; Type=" + properties.items[i].type + "; Property Value=" + properties.items[i].value);
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

