# Manage dropdown list content controls

Inserts, updates, and deletes dropdown list content controls.

This sample demonstrates how to insert, change, and delete dropdown list content controls.

```typescript
async function insertDropdownListContentControl() {
  // Places a dropdown list content control at the end of the selection.
  await Word.run(async (context) => {
    let selection = context.document.getSelection();
    selection.getRange(Word.RangeLocation.end).insertContentControl(Word.ContentControlType.dropDownList);
    await context.sync();

    console.log("Dropdown list content control inserted at the end of the selection.");
  });
}

async function addItemToDropdownListContentControl() {
  // Adds the provided list item to the first dropdown list content control in the selection.
  await Word.run(async (context) => {
    const listItemText = $("#item-to-add")
      .val()
      .toString()
      .trim();
    const selectedRange: Word.Range = context.document.getSelection();
    let selectedContentControl = selectedRange
      .getContentControls({
        types: [Word.ContentControlType.dropDownList]
      })
      .getFirstOrNullObject();
    selectedContentControl.load("id,dropDownListContentControl");
    await context.sync();

    if (selectedContentControl.isNullObject) {
      const parentContentControl: Word.ContentControl = selectedRange.parentContentControl;
      parentContentControl.load("id,type,dropDownListContentControl");
      await context.sync();

      if (parentContentControl.isNullObject || parentContentControl.type !== Word.ContentControlType.dropDownList) {
        console.warn("No dropdown list content control is currently selected.");
        return;
      } else {
        selectedContentControl = parentContentControl;
      }
    }

    selectedContentControl.dropDownListContentControl.addListItem(listItemText);
    await context.sync();

    console.log(`List item added to control with ID ${selectedContentControl.id}: ${listItemText}`);
  });
}

async function getListFromDropdownListContentControl() {
  // Gets the list items from the first dropdown list content control found in the selection.
  await Word.run(async (context) => {
    const selectedRange: Word.Range = context.document.getSelection();
    let selectedContentControl = selectedRange
      .getContentControls({
        types: [Word.ContentControlType.dropDownList]
      })
      .getFirstOrNullObject();
    selectedContentControl.load("id,dropDownListContentControl");
    await context.sync();

    if (selectedContentControl.isNullObject) {
      const parentContentControl: Word.ContentControl = selectedRange.parentContentControl;
      parentContentControl.load("id,type,dropDownListContentControl");
      await context.sync();

      if (parentContentControl.isNullObject || parentContentControl.type !== Word.ContentControlType.dropDownList) {
        console.warn("No dropdown list content control is currently selected.");
        return;
      } else {
        selectedContentControl = parentContentControl;
      }
    }

    let selectedDropdownList: Word.DropDownListContentControl = selectedContentControl.dropDownListContentControl;
    selectedDropdownList.listItems.load("items");
    await context.sync();

    const currentItems: Word.ContentControlListItemCollection = selectedDropdownList.listItems;
    console.log(`The list from the dropdown list content control with ID ${selectedContentControl.id}:`, currentItems);
  });
}

async function deleteItemFromDropdownListContentControl() {
  // Deletes the provided list item from the first dropdown list content control in the selection.
  await Word.run(async (context) => {
    const listItemText = $("#item-to-delete")
      .val()
      .toString()
      .trim();
    const selectedRange: Word.Range = context.document.getSelection();
    let selectedContentControl = selectedRange
      .getContentControls({
        types: [Word.ContentControlType.dropDownList]
      })
      .getFirstOrNullObject();
    selectedContentControl.load("id,dropDownListContentControl");
    await context.sync();

    if (selectedContentControl.isNullObject) {
      const parentContentControl: Word.ContentControl = selectedRange.parentContentControl;
      parentContentControl.load("id,type,dropDownListContentControl");
      await context.sync();

      if (parentContentControl.isNullObject || parentContentControl.type !== Word.ContentControlType.dropDownList) {
        console.warn("No dropdown list content control is currently selected.");
        return;
      } else {
        selectedContentControl = parentContentControl;
      }
    }

    let selectedDropdownList: Word.DropDownListContentControl = selectedContentControl.dropDownListContentControl;
    selectedDropdownList.listItems.load("items/*");
    await context.sync();

    let listItems: Word.ContentControlListItemCollection = selectedContentControl.dropDownListContentControl.listItems;
    let itemToDelete: Word.ContentControlListItem = listItems.items.find((item) => item.displayText === listItemText);
    if (!itemToDelete) {
      console.warn(`List item doesn't exist in control with ID ${selectedContentControl.id}: ${listItemText}`)
      return;
    }
    
    itemToDelete.delete();
    await context.sync();

    console.log(`List item deleted from control with ID ${selectedContentControl.id}: ${listItemText}`);
  });
}

async function deleteListFromDropdownListContentControl() {
  // Deletes the list items from first dropdown list content control found in the selection.
  await Word.run(async (context) => {
    const selectedRange: Word.Range = context.document.getSelection();
    let selectedContentControl = selectedRange
      .getContentControls({
        types: [Word.ContentControlType.dropDownList]
      })
      .getFirstOrNullObject();
    selectedContentControl.load("id,dropDownListContentControl");
    await context.sync();

    if (selectedContentControl.isNullObject) {
      const parentContentControl: Word.ContentControl = selectedRange.parentContentControl;
      parentContentControl.load("id,type,dropDownListContentControl");
      await context.sync();

      if (parentContentControl.isNullObject || parentContentControl.type !== Word.ContentControlType.dropDownList) {
        console.warn("No dropdown list content control is currently selected.");
        return;
      } else {
        selectedContentControl = parentContentControl;
      }
    }

    console.log(
      `About to delete the list from the dropdown list content control with ID ${selectedContentControl.id}`
    );
    selectedContentControl.dropDownListContentControl.deleteAllListItems();
    await context.sync();

    console.log("Deleted the list from the dropdown list content control.");
  });
}

async function deleteDropdownListContentControl() {
  // Deletes the first dropdown list content control found in the selection.
  await Word.run(async (context) => {
    const selectedRange: Word.Range = context.document.getSelection();
    let selectedContentControl = selectedRange
      .getContentControls({
        types: [Word.ContentControlType.dropDownList]
      })
      .getFirstOrNullObject();
    selectedContentControl.load("id");
    await context.sync();

    if (selectedContentControl.isNullObject) {
      const parentContentControl: Word.ContentControl = selectedRange.parentContentControl;
      parentContentControl.load("id,type");
      await context.sync();

      if (parentContentControl.isNullObject || parentContentControl.type !== Word.ContentControlType.dropDownList) {
        console.warn("No dropdown list content control is currently selected.");
        return;
      } else {
        selectedContentControl = parentContentControl;
      }
    }

    console.log(`About to delete dropdown list content control with ID ${selectedContentControl.id}`);
    selectedContentControl.delete(false);
    await context.sync();

    console.log("Deleted dropdown list content control.");
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
      console.warn("No dropdown list content control is currently selected.");
    } else {
      console.error(error);
    }
  }
}
```

