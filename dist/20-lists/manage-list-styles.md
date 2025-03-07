# Get list styles

This sample shows how to get the list styles in the current document.

This sample shows how to get the list styles in the current document.

```typescript
async function getCount() {
  // Gets the available list styles stored with the document.
  await Word.run(async (context) => {
    const styles: Word.StyleCollection = context.document.getStyles();
    const count = styles.getCount();

    // Load object to log properties and their values in the console.
    styles.load();
    await context.sync();

    for (let i = 0; i <= count.value; i++) {
      if (styles.items[i] && styles.items[i].type == "List") {
        console.log(`List style name: ${styles.items[i].nameLocal}`, styles.items[i]);
      }
    }
  });
}

async function getProperties() {
  // Gets the properties of the specified style.
  await Word.run(async (context) => {
    const styleName = $("#style-name-to-use").val() as string;
    if (styleName == "") {
      console.warn("Enter a style name to get properties.");
      return;
    }

    const style: Word.Style = context.document.getStyles().getByNameOrNullObject(styleName);
    style.load("type");
    await context.sync();

    if (style.isNullObject || style.type != Word.StyleType.list) {
      console.warn(`There's no existing style with the name '${styleName}'. Or this isn't a list style.`);
    } else {
      // Load objects to log properties and their values in the console.
      style.load();
      style.listTemplate.load();
      await context.sync();

      console.log(`Properties of the '${styleName}' style:`, style);

      const listLevels = style.listTemplate.listLevels;
      listLevels.load("items");
      await context.sync();

      console.log(`List levels of the '${styleName}' style:`, listLevels);
    }
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

