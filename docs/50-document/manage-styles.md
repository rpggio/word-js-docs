# Manage styles

This sample shows how to perform operations on the styles in the current document and how to add and delete custom styles.

This sample demonstrates how to manage styles.

```typescript
async function getCount() {
  // Gets the number of available styles stored with the document.
  await Word.run(async (context) => {
    const styles: Word.StyleCollection = context.document.getStyles();
    const count = styles.getCount();
    await context.sync();

    console.log(`Number of styles: ${count.value}`);
  });
}

async function addStyle() {
  // Adds a new style.
  await Word.run(async (context) => {
    const newStyleName = $("#new-style-name").val() as string;
    if (newStyleName == "") {
      console.warn("Enter a style name to add.");
      return;
    }

    const style: Word.Style = context.document.getStyles().getByNameOrNullObject(newStyleName);
    style.load();
    await context.sync();

    if (!style.isNullObject) {
      console.warn(
        `There's an existing style with the same name '${newStyleName}'! Please provide another style name.`
      );
      return;
    }

    const newStyleType = ($("#new-style-type").val() as unknown) as Word.StyleType;
    context.document.addStyle(newStyleName, newStyleType);
    await context.sync();

    console.log(newStyleName + " has been added to the style list.");
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
    style.load();
    await context.sync();

    if (style.isNullObject) {
      console.warn(`There's no existing style with the name '${styleName}'.`);
    } else {
      style.font.load();
      style.paragraphFormat.load();
      await context.sync();

      console.log(`Properties of the '${styleName}' style:`, style);
    }
  });
}

async function applyStyle() {
  // Applies the specified style to a paragraph.
  await Word.run(async (context) => {
    const styleName = $("#style-name-to-use").val() as string;
    if (styleName == "") {
      console.warn("Enter a style name to apply.");
      return;
    }

    const style: Word.Style = context.document.getStyles().getByNameOrNullObject(styleName);
    style.load();
    await context.sync();

    if (style.isNullObject) {
      console.warn(`There's no existing style with the name '${styleName}'.`);
    } else if (style.type != Word.StyleType.paragraph) {
      console.log(`The '${styleName}' style isn't a paragraph style.`);
    } else {
      const body: Word.Body = context.document.body;
      body.clear();
      body.insertParagraph(
        "Do you want to create a solution that extends the functionality of Word? You can use the Office Add-ins platform to extend Word clients running on the web, on a Windows desktop, or on a Mac.",
        "Start"
      );
      const paragraph: Word.Paragraph = body.paragraphs.getFirst();
      paragraph.style = style.nameLocal;
      console.log(`'${styleName}' style applied to first paragraph.`);
    }
  });
}

async function setFontProperties() {
  // Updates font properties (e.g., color, size) of the specified style.
  await Word.run(async (context) => {
    const styleName = $("#style-name").val() as string;
    if (styleName == "") {
      console.warn("Enter a style name to update font properties.");
      return;
    }

    const style: Word.Style = context.document.getStyles().getByNameOrNullObject(styleName);
    style.load();
    await context.sync();

    if (style.isNullObject) {
      console.warn(`There's no existing style with the name '${styleName}'.`);
    } else {
      const font: Word.Font = style.font;
      font.color = "#FF0000";
      font.size = 20;
      console.log(`Successfully updated font properties of the '${styleName}' style.`);
    }
  });
}

async function setParagraphFormat() {
  // Sets certain aspects of the specified style's paragraph format e.g., the left indent size and the alignment.
  await Word.run(async (context) => {
    const styleName = $("#style-name").val() as string;
    if (styleName == "") {
      console.warn("Enter a style name to update its paragraph format.");
      return;
    }

    const style: Word.Style = context.document.getStyles().getByNameOrNullObject(styleName);
    style.load();
    await context.sync();

    if (style.isNullObject) {
      console.warn(`There's no existing style with the name '${styleName}'.`);
    } else {
      style.paragraphFormat.leftIndent = 30;
      style.paragraphFormat.alignment = Word.Alignment.centered;
      console.log(`Successfully the paragraph format of the '${styleName}' style.`);
    }
  });
}

async function setBorderProperties() {
  // Updates border properties (e.g., type, width, color) of the specified style.
  await Word.run(async (context) => {
    const styleName = $("#style-name").val() as string;
    if (styleName == "") {
      console.warn("Enter a style name to update border properties.");
      return;
    }

    const style: Word.Style = context.document.getStyles().getByNameOrNullObject(styleName);
    style.load();
    await context.sync();

    if (style.isNullObject) {
      console.warn(`There's no existing style with the name '${styleName}'.`);
    } else {
      const borders: Word.BorderCollection = style.borders;
      borders.load("items");
      await context.sync();

      borders.outsideBorderType = Word.BorderType.dashed;
      borders.outsideBorderWidth = Word.BorderWidth.pt025;
      borders.outsideBorderColor = "green";
      console.log("Updated outside borders.");
    }
  });
}

async function setShadingProperties() {
  // Updates shading properties (e.g., texture, pattern colors) of the specified style.
  await Word.run(async (context) => {
    const styleName = $("#style-name").val() as string;
    if (styleName == "") {
      console.warn("Enter a style name to update shading properties.");
      return;
    }

    const style: Word.Style = context.document.getStyles().getByNameOrNullObject(styleName);
    style.load();
    await context.sync();

    if (style.isNullObject) {
      console.warn(`There's no existing style with the name '${styleName}'.`);
    } else {
      const shading: Word.Shading = style.shading;
      shading.load();
      await context.sync();

      shading.backgroundPatternColor = "blue";
      shading.foregroundPatternColor = "yellow";
      shading.texture = Word.ShadingTextureType.darkTrellis;

      console.log("Updated shading.");
    }
  });
}

async function deleteStyle() {
  // Deletes the custom style.
  await Word.run(async (context) => {
    const styleName = $("#style-name-to-delete").val() as string;
    if (styleName == "") {
      console.warn("Enter a style name to delete.");
      return;
    }

    const style: Word.Style = context.document.getStyles().getByNameOrNullObject(styleName);
    style.load();
    await context.sync();

    if (style.isNullObject) {
      console.warn(`There's no existing style with the name '${styleName}'.`);
    } else {
      style.delete();
      console.log(`Successfully deleted custom style '${styleName}'.`);
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

