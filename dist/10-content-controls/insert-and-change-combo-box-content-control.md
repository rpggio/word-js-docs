# Manage combo box content controls

Inserts, updates, and deletes combo box content controls.

This sample demonstrates how to insert, change, and delete combo box content controls.

```typescript
async function insertComboBoxContentControl() {
  // Places a combo box content control at the end of the selection.
  await Word.run(async (context) => {
    let selection = context.document.getSelection();
    selection.getRange(Word.RangeLocation.end).insertContentControl(Word.ContentControlType.comboBox);
    await context.sync();

    console.log("Combo box content control inserted at the end of the selection.");
  });
}

async function addItemToComboBoxContentControl() {
  // Adds the provided list item to the first combo box content control in the selection.
  await Word.run(async (context) => {
    const listItemText = $("#item-to-add")
      .val()
      .toString()
      .trim();
    const selectedRange: Word.Range = context.document.getSelection();
    let selectedContentControl = selectedRange
      .getContentControls({
        types: [Word.ContentControlType.comboBox]
      })
      .getFirstOrNullObject();
    selectedContentControl.load("id,comboBoxContentControl");
    await context.sync();

    if (selectedContentControl.isNullObject) {
      const parentContentControl: Word.ContentControl = selectedRange.parentContentControl;
      parentContentControl.load("id,type,comboBoxContentControl");
      await context.sync();

      if (parentContentControl.isNullObject || parentContentControl.type !== Word.ContentControlType.comboBox) {
        console.warn("No combo box content control is currently selected.");
        return;
      } else {
        selectedContentControl = parentContentControl;
      }
    }

    selectedContentControl.comboBoxContentControl.addListItem(listItemText);
    await context.sync();

    console.log(`List item added to control with ID ${selectedContentControl.id}: ${listItemText}`);
  });
}

async function getListFromComboBoxContentControl() {
  // Gets the list items from the first combo box content control found in the selection.
  await Word.run(async (context) => {
    const selectedRange: Word.Range = context.document.getSelection();
    let selectedContentControl = selectedRange
      .getContentControls({
        types: [Word.ContentControlType.comboBox]
      })
      .getFirstOrNullObject();
    selectedContentControl.load("id,comboBoxContentControl");
    await context.sync();

    if (selectedContentControl.isNullObject) {
      const parentContentControl: Word.ContentControl = selectedRange.parentContentControl;
      parentContentControl.load("id,type,comboBoxContentControl");
      await context.sync();

      if (parentContentControl.isNullObject || parentContentControl.type !== Word.ContentControlType.comboBox) {
        console.warn("No combo box content control is currently selected.");
        return;
      } else {
        selectedContentControl = parentContentControl;
      }
    }

    let selectedComboBox: Word.ComboBoxContentControl = selectedContentControl.comboBoxContentControl;
    selectedComboBox.listItems.load("items");
    await context.sync();

    const currentItems: Word.ContentControlListItemCollection = selectedComboBox.listItems;
    console.log(`The list from the combo box content control with ID ${selectedContentControl.id}:`, currentItems);
  });
}

async function deleteItemFromComboBoxContentControl() {
  // Deletes the provided list item from the first combo box content control in the selection.
  await Word.run(async (context) => {
    const listItemText = $("#item-to-delete")
      .val()
      .toString()
      .trim();
    const selectedRange: Word.Range = context.document.getSelection();
    let selectedContentControl = selectedRange
      .getContentControls({
        types: [Word.ContentControlType.comboBox]
      })
      .getFirstOrNullObject();
    selectedContentControl.load("id,comboBoxContentControl");
    await context.sync();

    if (selectedContentControl.isNullObject) {
      const parentContentControl: Word.ContentControl = selectedRange.parentContentControl;
      parentContentControl.load("id,type,comboBoxContentControl");
      await context.sync();

      if (parentContentControl.isNullObject || parentContentControl.type !== Word.ContentControlType.comboBox) {
        console.warn("No combo box content control is currently selected.");
        return;
      } else {
        selectedContentControl = parentContentControl;
      }
    }

    let selectedComboBox: Word.ComboBoxContentControl = selectedContentControl.comboBoxContentControl;
    selectedComboBox.listItems.load("items/*");
    await context.sync();

    let listItems: Word.ContentControlListItemCollection = selectedContentControl.comboBoxContentControl.listItems;
    let itemToDelete: Word.ContentControlListItem = listItems.items.find((item) => item.displayText === listItemText);
    if (!itemToDelete) {
      console.warn(`List item doesn't exist in control with ID ${selectedContentControl.id}: ${listItemText}`);
      return;
    }

    itemToDelete.delete();
    await context.sync();

    console.log(`List item deleted from control with ID ${selectedContentControl.id}: ${listItemText}`);
  });
}

async function deleteListFromComboBoxContentControl() {
  // Deletes the list items from first combo box content control found in the selection.
  await Word.run(async (context) => {
    const selectedRange: Word.Range = context.document.getSelection();
    let selectedContentControl = selectedRange
      .getContentControls({
        types: [Word.ContentControlType.comboBox]
      })
      .getFirstOrNullObject();
    selectedContentControl.load("id,comboBoxContentControl");
    await context.sync();

    if (selectedContentControl.isNullObject) {
      const parentContentControl: Word.ContentControl = selectedRange.parentContentControl;
      parentContentControl.load("id,type,comboBoxContentControl");
      await context.sync();

      if (parentContentControl.isNullObject || parentContentControl.type !== Word.ContentControlType.comboBox) {
        console.warn("No combo box content control is currently selected.");
        return;
      } else {
        selectedContentControl = parentContentControl;
      }
    }

    console.log(`About to delete the list from the combo box content control with ID ${selectedContentControl.id}`);
    selectedContentControl.comboBoxContentControl.deleteAllListItems();
    await context.sync();

    console.log("Deleted the list from the combo box content control.");
  });
}

async function deleteComboBoxContentControl() {
  // Deletes the first combo box content control found in the selection.
  await Word.run(async (context) => {
    const selectedRange: Word.Range = context.document.getSelection();
    let selectedContentControl = selectedRange
      .getContentControls({
        types: [Word.ContentControlType.comboBox]
      })
      .getFirstOrNullObject();
    selectedContentControl.load("id");
    await context.sync();

    if (selectedContentControl.isNullObject) {
      const parentContentControl: Word.ContentControl = selectedRange.parentContentControl;
      parentContentControl.load("id,type");
      await context.sync();

      if (parentContentControl.isNullObject || parentContentControl.type !== Word.ContentControlType.comboBox) {
        console.warn("No combo box content control is currently selected.");
        return;
      } else {
        selectedContentControl = parentContentControl;
      }
    }

    console.log(`About to delete combo box content control with ID ${selectedContentControl.id}`);
    selectedContentControl.delete(false);
    await context.sync();

    console.log("Deleted combo box content control.");
  });
}

async function setup() {
  await Word.run(async (context) => {
    const body: Word.Body = context.document.body;
    body.clear();
    body.insertParagraph("One more paragraph.", "Start");
    body.insertParagraph("Inserting another paragraph.", "Start");
    body.insertParagraph(
      "Do you want to create a solution that extends the functionality of Word? You can use the Office Add-ins platform to extend Word clients running on the web, on a Windows desktop, or on a Mac.",
      "Start"
    );
    body.paragraphs
      .getLast()
      .insertText(
        "Use add-in commands to extend the Word UI and launch task panes that run JavaScript that interacts with the content in a Word document. Any code that you can run in a browser can run in a Word add-in. Add-ins that interact with content in a Word document create requests to act on Word objects and synchronize object state.",
        "Replace"
      );
  });
}

// Default helper for invoking an action and handling errors.
async function tryCatch(callback) {
  try {
    await callback();
  } catch (error) {
    // Note: In a production add-in, you'd want to notify the user through your add-in's UI.
    if (error.code === Word.ErrorCodes.itemNotFound) {
      console.warn("No combo box content control is currently selected.");
    } else {
      console.error(error);
    }
  }
}
```

