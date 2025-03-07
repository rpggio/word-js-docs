# Table formatting

Gets the formatting details of a table, a table row, and a table cell, including borders, alignment, and cell padding.

This sample shows how to get various formatting details about a table, a table row, and a table cell, including
    borders, alignment, and cell padding.

```typescript
async function getTableAlignment() {
  // Gets alignment details about the first table in the document.
  await Word.run(async (context) => {
    const firstTable: Word.Table = context.document.body.tables.getFirst();
    firstTable.load(["alignment", "horizontalAlignment", "verticalAlignment"]);
    await context.sync();

    console.log(`Details about the alignment of the first table:`, `- Alignment of the table within the containing page column: ${firstTable.alignment}`, `- Horizontal alignment of every cell in the table: ${firstTable.horizontalAlignment}`, `- Vertical alignment of every cell in the table: ${firstTable.verticalAlignment}`);
  });
}

async function getTableBorder() {
  // Gets border details about the first table in the document.
  await Word.run(async (context) => {
    const firstTable: Word.Table = context.document.body.tables.getFirst();
    const borderLocation = Word.BorderLocation.top;
    const border: Word.TableBorder = firstTable.getBorder(borderLocation);
    border.load(["type", "color", "width"]);
    await context.sync();

    console.log(`Details about the ${borderLocation} border of the first table:`, `- Color: ${border.color}`, `- Type: ${border.type}`, `- Width: ${border.width} points`);
  });
}

async function getTableCellPadding() {
  // Gets cell padding details about the first table in the document.
  await Word.run(async (context) => {
    const firstTable: Word.Table = context.document.body.tables.getFirst();
    const cellPaddingLocation = Word.CellPaddingLocation.right;
    const cellPadding = firstTable.getCellPadding(cellPaddingLocation);
    await context.sync();

    console.log(
      `Cell padding details about the ${cellPaddingLocation} border of the first table: ${cellPadding.value} points`
    );
  });
}

async function getTableRowAlignment() {
  // Gets content alignment details about the first row of the first table in the document.
  await Word.run(async (context) => {
    const firstTable: Word.Table = context.document.body.tables.getFirst();
    const firstTableRow: Word.TableRow = firstTable.rows.getFirst();
    firstTableRow.load(["horizontalAlignment", "verticalAlignment"]);
    await context.sync();

    console.log(`Details about the alignment of the first table's first row:`, `- Horizontal alignment of every cell in the row: ${firstTableRow.horizontalAlignment}`, `- Vertical alignment of every cell in the row: ${firstTableRow.verticalAlignment}`);
  });
}

async function getTableRowBorder() {
  // Gets border details about the first row of the first table in the document.
  await Word.run(async (context) => {
    const firstTable: Word.Table = context.document.body.tables.getFirst();
    const firstTableRow: Word.TableRow = firstTable.rows.getFirst();
    const borderLocation = Word.BorderLocation.bottom;
    const border: Word.TableBorder = firstTableRow.getBorder(borderLocation);
    border.load(["type", "color", "width"]);
    await context.sync();

    console.log(`Details about the ${borderLocation} border of the first table's first row:`, `- Color: ${border.color}`, `- Type: ${border.type}`, `- Width: ${border.width} points`);
  });
}

async function getTableRowCellPadding() {
  // Gets cell padding details about the first row of the first table in the document.
  await Word.run(async (context) => {
    const firstTable: Word.Table = context.document.body.tables.getFirst();
    const firstTableRow: Word.TableRow = firstTable.rows.getFirst();
    const cellPaddingLocation = Word.CellPaddingLocation.bottom;
    const cellPadding = firstTableRow.getCellPadding(cellPaddingLocation);
    await context.sync();

    console.log(
      `Cell padding details about the ${cellPaddingLocation} border of the first table's first row: ${cellPadding.value} points`
    );
  });
}

async function getTableCellAlignment() {
  // Gets content alignment details about the first cell of the first table in the document.
  await Word.run(async (context) => {
    const firstTable: Word.Table = context.document.body.tables.getFirst();
    const firstTableRow: Word.TableRow = firstTable.rows.getFirst();
    const firstCell: Word.TableCell = firstTableRow.cells.getFirst();
    firstCell.load(["horizontalAlignment", "verticalAlignment"]);
    await context.sync();

    console.log(`Details about the alignment of the first table's first cell:`, `- Horizontal alignment of the cell's content: ${firstCell.horizontalAlignment}`, `- Vertical alignment of the cell's content: ${firstCell.verticalAlignment}`);
  });
}

async function getTableCellBorder() {
  // Gets border details about the first of the first table in the document.
  await Word.run(async (context) => {
    const firstTable: Word.Table = context.document.body.tables.getFirst();
    const firstCell: Word.TableCell = firstTable.getCell(0, 0);
    const borderLocation = "Left";
    const border: Word.TableBorder = firstCell.getBorder(borderLocation);
    border.load(["type", "color", "width"]);
    await context.sync();

    console.log(`Details about the ${borderLocation} border of the first table's first cell:`, `- Color: ${border.color}`, `- Type: ${border.type}`, `- Width: ${border.width} points`);
  });
}

async function getTableCellCellPadding() {
  // Gets cell padding details about the first cell of the first table in the document.
  await Word.run(async (context) => {
    const firstTable: Word.Table = context.document.body.tables.getFirst();
    const firstCell: Word.TableCell = firstTable.getCell(0, 0);
    const cellPaddingLocation = "Left";
    const cellPadding = firstCell.getCellPadding(cellPaddingLocation);
    await context.sync();

    console.log(
      `Cell padding details about the ${cellPaddingLocation} border of the first table's first cell: ${cellPadding.value} points`
    );
  });
}

async function insertTable() {
  await Word.run(async (context) => {
    // Use a two-dimensional array to hold the initial table values.
    const data = [
      ["Tokyo", "Beijing", "Seattle"],
      ["Apple", "Orange", "Pineapple"]
    ];
    const table = context.document.body.insertTable(2, 3, "Start", data);
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

