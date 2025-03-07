# Manage checkbox content controls

Inserts, updates, retrieves, and deletes checkbox content controls.

This sample demonstrates how to insert, change, and delete checkbox content controls.

```typescript
async function insertCheckboxContentControls() {
  // Traverses each paragraph of the document and places a checkbox content control at the beginning of each.
  await Word.run(async (context) => {
    let paragraphs = context.document.body.paragraphs;
    paragraphs.load("$none"); // Don't need any properties; just start each paragraph with a content control.

    await context.sync();

    for (let i = 0; i < paragraphs.items.length; i++) {
      let contentControl = paragraphs.items[i]
        .getRange(Word.RangeLocation.start)
        .insertContentControl(Word.ContentControlType.checkBox);
    }
    console.log("Checkbox content controls inserted: " + paragraphs.items.length);

    await context.sync();
  });
}

async function toggleCheckboxContentControl() {
  // Toggles the isChecked property of the first checkbox content control found in the selection.
  await Word.run(async (context) => {
    const selectedRange: Word.Range = context.document.getSelection();
    let selectedContentControl = selectedRange
      .getContentControls({
        types: [Word.ContentControlType.checkBox]
      })
      .getFirstOrNullObject();
    selectedContentControl.load("id,checkboxContentControl/isChecked");

    await context.sync();

    if (selectedContentControl.isNullObject) {
      const parentContentControl: Word.ContentControl = selectedRange.parentContentControl;
      parentContentControl.load("id,type,checkboxContentControl/isChecked");
      await context.sync();

      if (parentContentControl.isNullObject || parentContentControl.type !== Word.ContentControlType.checkBox) {
        console.warn("No checkbox content control is currently selected.");
        return;
      } else {
        selectedContentControl = parentContentControl;
      }
    }

    const isCheckedBefore = selectedContentControl.checkboxContentControl.isChecked;
    console.log("isChecked state before:", `id: ${selectedContentControl.id} ... isChecked: ${isCheckedBefore}`);
    selectedContentControl.checkboxContentControl.isChecked = !isCheckedBefore;
    selectedContentControl.load("id,checkboxContentControl/isChecked");
    await context.sync();

    console.log(
      "isChecked state after:",
      `id: ${selectedContentControl.id} ... isChecked: ${selectedContentControl.checkboxContentControl.isChecked}`
    );
  });
}

async function toggleCheckboxContentControls() {
  // Toggles the isChecked property on all checkbox content controls.
  await Word.run(async (context) => {
    let contentControls = context.document.getContentControls({
      types: [Word.ContentControlType.checkBox]
    });
    contentControls.load("items");

    await context.sync();

    const length = contentControls.items.length;
    console.log(`Number of checkbox content controls: ${length}`);

    if (length <= 0) {
      return;
    }

    const checkboxContentControls = [];
    for (let i = 0; i < length; i++) {
      let contentControl = contentControls.items[i];
      contentControl.load("id,checkboxContentControl/isChecked");
      checkboxContentControls.push(contentControl);
    }

    await context.sync();

    console.log("isChecked state before:");
    const updatedCheckboxContentControls = [];
    for (let i = 0; i < checkboxContentControls.length; i++) {
      const currentCheckboxContentControl = checkboxContentControls[i];
      const isCheckedBefore = currentCheckboxContentControl.checkboxContentControl.isChecked;
      console.log(`id: ${currentCheckboxContentControl.id} ... isChecked: ${isCheckedBefore}`);

      currentCheckboxContentControl.checkboxContentControl.isChecked = !isCheckedBefore;
      currentCheckboxContentControl.load("id,checkboxContentControl/isChecked");
      updatedCheckboxContentControls.push(currentCheckboxContentControl);
    }

    await context.sync();

    console.log("isChecked state after:");
    for (let i = 0; i < updatedCheckboxContentControls.length; i++) {
      const currentCheckboxContentControl = updatedCheckboxContentControls[i];
      console.log(
        `id: ${currentCheckboxContentControl.id} ... isChecked: ${currentCheckboxContentControl.checkboxContentControl.isChecked}`
      );
    }
  });
}

async function deleteCheckboxContentControl() {
  // Deletes the first checkbox content control found in the selection.
  await Word.run(async (context) => {
    const selectedRange: Word.Range = context.document.getSelection();
    let selectedContentControl = selectedRange
      .getContentControls({
        types: [Word.ContentControlType.checkBox]
      })
      .getFirstOrNullObject();
    selectedContentControl.load("id");

    await context.sync();

    if (selectedContentControl.isNullObject) {
      const parentContentControl: Word.ContentControl = selectedRange.parentContentControl;
      parentContentControl.load("id,type");
      await context.sync();

      if (parentContentControl.isNullObject || parentContentControl.type !== Word.ContentControlType.checkBox) {
        console.warn("No checkbox content control is currently selected.");
        return;
      } else {
        selectedContentControl = parentContentControl;
      }
    }

    console.log(`About to delete checkbox content control with id: ${selectedContentControl.id}`);
    selectedContentControl.delete(false);
    await context.sync();

    console.log("Deleted checkbox content control.");
  });
}

async function setup() {
  await Word.run(async (context) => {
    const body: Word.Body = context.document.body;
    body.clear();
    body.insertParagraph("Task 3", "Start");
    body.insertParagraph("Task 2", "Start");
    body.insertParagraph("Task 1", "Start");
    body.paragraphs.getLast().insertText("Task 4", "Replace");
  });
}

// Default helper for invoking an action and handling errors.
async function tryCatch(callback) {
  try {
    await callback();
  } catch (error) {
    // Note: In a production add-in, you'd want to notify the user through your add-in's UI.
    if (error.code === Word.ErrorCodes.itemNotFound) {
      console.warn("No checkbox content control is currently selected.");
    } else {
      console.error(error);
    }
  }
}
```

