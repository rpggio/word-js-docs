# Create and access a table

Creates a table and accesses a specific cell.

This sample demonstrates how to get a cell from a table.

```typescript
async function getTableCell() {
  // Gets the content of the first cell in the first table.
  await Word.run(async (context) => {
    const firstCell: Word.Body = context.document.body.tables.getFirst().getCell(0, 0).body;
    firstCell.load("text");

    await context.sync();
    console.log("First cell's text is: " + firstCell.text);
  });
}

async function insertTable() {
  await Word.run(async (context) => {
    // Use a two-dimensional array to hold the initial table values.
    const data = [
      ["Tokyo", "Beijing", "Seattle"],
      ["Apple", "Orange", "Pineapple"]
    ];
    const table: Word.Table = context.document.body.insertTable(2, 3, "Start", data);
    table.styleBuiltIn = Word.BuiltInStyleName.gridTable5Dark_Accent2;
    table.styleFirstColumn = false;

    await context.sync();
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

