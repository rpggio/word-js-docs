# Manage fields

This sample shows how to perform basic operations on fields, including insert, get, and delete.

This sample demonstrates how to manage fields, which are placeholders for dynamic data in your document. To learn
        more about fields, refer to Insert, edit, and view
            fields in Word.

```typescript
async function rangeInsertDateField() {
  // Inserts a Date field before selection.
  await Word.run(async (context) => {
    const range: Word.Range = context.document.getSelection().getRange();

    const field: Word.Field = range.insertField(Word.InsertLocation.before, Word.FieldType.date, '\\@ "M/d/yyyy h:mm am/pm"', true);

    field.load("result,code");
    await context.sync();

    if (field.isNullObject) {
      console.log("There are no fields in this document.");
    } else {
      console.log("Code of the field: " + field.code, "Result of the field: " + JSON.stringify(field.result));
    }
  });
}

async function getFirstField() {
  // Gets the first field in the document.
  await Word.run(async (context) => {
    const field: Word.Field = context.document.body.fields.getFirstOrNullObject();
    field.load(["code", "result", "locked", "type", "data", "kind"]);

    await context.sync();

    if (field.isNullObject) {
      console.log("This document has no fields.");
    } else {
      console.log("Code of first field: " + field.code, "Result of first field: " + JSON.stringify(field.result), "Type of first field: " + field.type, "Is the first field locked? " + field.locked, "Kind of the first field: " + field.kind);
    }
  });
}

async function getParentBodyOfFirstField() {
  // Gets the parent body of the first field in the document.
  await Word.run(async (context) => {
    const field: Word.Field = context.document.body.fields.getFirstOrNullObject();
    field.load("parentBody/text");

    await context.sync();

    if (field.isNullObject) {
      console.log("This document has no fields.");
    } else {
      const parentBody: Word.Body = field.parentBody;
      console.log("Text of first field's parent body: " + JSON.stringify(parentBody.text));
    }
  });
}

async function getAllFields() {
  // Gets all fields in the document body.
  await Word.run(async (context) => {
    const fields: Word.FieldCollection = context.document.body.fields.load("items");

    await context.sync();

    if (fields.items.length === 0) {
      console.log("No fields in this document.");
    } else {
      fields.load(["code", "result"]);
      await context.sync();

      for (let i = 0; i < fields.items.length; i++) {
        console.log(`Field ${i + 1}'s code: ${fields.items[i].code}`, `Field ${i + 1}'s result: ${JSON.stringify(fields.items[i].result)}`);
      }
    }
  });
}

async function getSelectedFieldAndUpdate() {
  // Gets and updates the first field in the selection.
  await Word.run(async (context) => {
    let field = context.document.getSelection().fields.getFirstOrNullObject();
    field.load(["code", "result", "type", "locked"]);

    await context.sync();

    if (field.isNullObject) {
      console.log("No field in selection.");
    } else {
      console.log("Before updating:", "Code of first field: " + field.code, "Result of first field: " + JSON.stringify(field.result));

      field.updateResult();
      field.select();
      await context.sync();

      field.load(["code", "result"]);
      await context.sync();

      console.log("After updating:", "Code of first field: " + field.code, "Result of first field: " + JSON.stringify(field.result));
    }
  });
}

async function getFieldAndLockOrUnlock() {
  // Gets the first field in the selection and toggles between setting it to locked or unlocked.
  await Word.run(async (context) => {
    let field = context.document.getSelection().fields.getFirstOrNullObject();
    field.load(["code", "result", "type", "locked"]);
    await context.sync();

    if (field.isNullObject) {
      console.log("The selection has no fields.");
    } else {
      console.log(`The first field in the selection is currently ${field.locked ? "locked" : "unlocked"}.`);
      field.locked = !field.locked;
      await context.sync();

      console.log(`The first field in the selection is now ${field.locked ? "locked" : "unlocked"}.`);
    }
  });
}

async function deleteFirstField() {
  // Deletes the first field in the document.
  await Word.run(async (context) => {
    const field: Word.Field = context.document.body.fields.getFirstOrNullObject();
    field.load();

    await context.sync();

    if (field.isNullObject) {
      console.log("This document has no fields.");
    } else {
      field.delete();
      await context.sync();

      console.log("The first field in the document was deleted.");
    }
  });
}

async function setup() {
  await Word.run(async (context) => {
    const body: Word.Body = context.document.body;
    body.clear();
    body.insertParagraph(
      "Do you want to create a solution that extends the functionality of Word? You can use the Office Add-ins platform to extend Word clients running on the web, on a Windows desktop, or on a Mac.",
      "Start"
    );
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

