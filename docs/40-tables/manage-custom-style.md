# Manage custom table style

Shows how to manage primarily margins and alignments of a custom table style in the current document.

This sample demonstrates how to manage a custom table style and use Document.importStylesFromJson.

    **Important**: Some TableStyle properties are currently in preview. If this snippet doesn't work, try using Word
        on a different platform.

```typescript
async function addStyle() {
  // Adds a new table style.
  const newStyleName = $("#new-style-name").val() as string;
  if (newStyleName == "") {
    console.warn("Enter a style name to add.");
    return;
  }

  await Word.run(async (context) => {
    const style: Word.Style = context.document.getStyles().getByNameOrNullObject(newStyleName);
    style.load();
    await context.sync();

    if (!style.isNullObject) {
      console.warn(
        `There's an existing style with the same name '${newStyleName}'! Please provide another style name.`
      );
      return;
    }

    context.document.addStyle(newStyleName, Word.StyleType.table);
    await context.sync();

    console.log(newStyleName + " has been added to the style list.");
  });
}

async function applyStyle() {
  // Applies the specified style to a new table.
  const styleName = $("#style-name").val() as string;
  if (styleName == "") {
    console.warn("Enter a style name to apply.");
    return;
  }

  await Word.run(async (context) => {
    const style: Word.Style = context.document.getStyles().getByNameOrNullObject(styleName);
    style.load();
    await context.sync();

    if (style.isNullObject) {
      console.warn(`There's no existing style with the name '${styleName}'.`);
    } else if (style.type != Word.StyleType.table) {
      console.warn(`The '${styleName}' style isn't a table style.`);
    } else {
      const body: Word.Body = context.document.body;
      body.clear();
      const data = [
        ["Tokyo", "Beijing", "Seattle"],
        ["Apple", "Orange", "Pineapple"]
      ];
      const table: Word.Table = body.insertTable(2, 3, "Start", data);
      table.style = style.nameLocal;
      table.styleFirstColumn = false;
      await context.sync();

      console.log(`'${styleName}' style applied to first table.`, style);
    }
  });
}

async function getTableStyle() {
  // Gets the table style properties and displays them in the form.
  const styleName = $("#style-name")
    .val()
    .toString();
  if (styleName == "") {
    console.warn("Please input a table style name.");
    return;
  }

  await Word.run(async (context) => {
    const tableStyle: Word.TableStyle = context.document.getStyles().getByName(styleName).tableStyle;
    tableStyle.load();
    await context.sync();

    if (tableStyle.isNullObject) {
      console.warn(`There's no existing table style with the name '${styleName}'.`);
      return;
    }

    console.log(tableStyle);
  });
}

async function setAlignment() {
  // Sets the table alignment.
  const styleName = $("#style-name")
    .val()
    .toString();
  if (styleName == "") {
    console.warn("Please input a table style name.");
    return;
  }

  await Word.run(async (context) => {
    const alignment = $("#alignment")
      .val()
      .toString();
    const tableStyle: Word.TableStyle = context.document.getStyles().getByName(styleName).tableStyle;
    tableStyle.alignment = alignment as Word.Alignment;
    await context.sync();

    tableStyle.load();
    await context.sync();
    console.log("Alignment: " + tableStyle.alignment);
  });
}

async function setAllowBreakAcrossPage() {
  // Sets the allowBreakAcrossPage property.
  const styleName = $("#style-name")
    .val()
    .toString();
  if (styleName == "") {
    console.warn("Please input a table style name.");
    return;
  }

  await Word.run(async (context) => {
    const allowBreakAcrossPage = $("#allow-break-across-page").val() as string;
    const tableStyle: Word.TableStyle = context.document.getStyles().getByName(styleName).tableStyle;
    tableStyle.allowBreakAcrossPage = allowBreakAcrossPage === "true";
    await context.sync();

    tableStyle.load();
    await context.sync();
    console.log("allowBreakAcrossPage: " + tableStyle.allowBreakAcrossPage);
  });
}

async function setTopCellMargin() {
  // Sets the top cell margin.
  const styleName = $("#style-name")
    .val()
    .toString();
  if (styleName == "") {
    console.warn("Please input a table style name.");
    return;
  }

  await Word.run(async (context) => {
    const topCellMargin = Number(
        .val()
        .toString()
    );
    const tableStyle: Word.TableStyle = context.document.getStyles().getByName(styleName).tableStyle;
    tableStyle.topCellMargin = topCellMargin;
    await context.sync();

    tableStyle.load();
    await context.sync();
    console.log("Top cell margin: " + tableStyle.topCellMargin);
  });
}

async function setBottomCellMargin() {
  // Sets the bottom cell margin.
  const styleName = $("#style-name")
    .val()
    .toString();
  if (styleName == "") {
    console.warn("Please input a table style name.");
    return;
  }

  await Word.run(async (context) => {
    const bottomCellMargin = Number(
        .val()
        .toString()
    );
    const tableStyle: Word.TableStyle = context.document.getStyles().getByName(styleName).tableStyle;
    tableStyle.bottomCellMargin = bottomCellMargin;
    await context.sync();

    tableStyle.load();
    await context.sync();
    console.log("Bottom cell margin: " + tableStyle.bottomCellMargin);
  });
}

async function setLeftCellMargin() {
  // Sets the left cell margin.
  const styleName = $("#style-name")
    .val()
    .toString();
  if (styleName == "") {
    console.warn("Please input a table style name.");
    return;
  }

  await Word.run(async (context) => {
    const leftCellMargin = Number(
        .val()
        .toString()
    );
    const tableStyle: Word.TableStyle = context.document.getStyles().getByName(styleName).tableStyle;
    tableStyle.leftCellMargin = leftCellMargin;
    await context.sync();

    tableStyle.load();
    await context.sync();
    console.log("Left cell margin: " + tableStyle.leftCellMargin);
  });
}

async function setRightCellMargin() {
  // Sets the right cell margin.
  const styleName = $("#style-name")
    .val()
    .toString();
  if (styleName == "") {
    console.warn("Please input a table style name.");
    return;
  }

  await Word.run(async (context) => {
    const rightCellMargin = Number(
        .val()
        .toString()
    );
    const tableStyle: Word.TableStyle = context.document.getStyles().getByName(styleName).tableStyle;
    tableStyle.rightCellMargin = rightCellMargin;
    await context.sync();

    tableStyle.load();
    await context.sync();
    console.log("Right cell margin: " + tableStyle.rightCellMargin);
  });
}

async function setCellSpacing() {
  // Sets the cell spacing.
  const styleName = $("#style-name")
    .val()
    .toString();
  if (styleName == "") {
    console.warn("Please input a table style name.");
    return;
  }

  await Word.run(async (context) => {
    const cellSpacing = Number(
        .val()
        .toString()
    );
    const tableStyle: Word.TableStyle = context.document.getStyles().getByName(styleName).tableStyle;
    tableStyle.cellSpacing = cellSpacing;
    await context.sync();

    tableStyle.load();
    await context.sync();
    console.log("Cell spacing: " + tableStyle.cellSpacing);
  });
}

async function deleteStyle() {
  // Deletes the custom style.
  const styleName = $("#style-name-to-delete").val() as string;
  if (styleName == "") {
    console.warn("Enter a style name to delete.");
    return;
  }

  await Word.run(async (context) => {
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

async function importStylesFromJson() {
  // Imports styles from JSON.
  await Word.run(async (context) => {
    const str =
      '{"styles":[{"baseStyle":"Default Paragraph Font","builtIn":false,"inUse":true,"linked":false,"nameLocal":"NewCharStyle","priority":2,"quickStyle":true,"type":"Character","unhideWhenUsed":false,"visibility":false,"paragraphFormat":null,"font":{"name":"DengXian Light","size":16.0,"bold":true,"italic":false,"color":"#F1A983","underline":"None","subscript":false,"superscript":true,"strikeThrough":true,"doubleStrikeThrough":false,"highlightColor":null,"hidden":false},"shading":{"backgroundPatternColor":"#FF0000"}},{"baseStyle":"Normal","builtIn":false,"inUse":true,"linked":false,"nextParagraphStyle":"NewParaStyle","nameLocal":"NewParaStyle","priority":1,"quickStyle":true,"type":"Paragraph","unhideWhenUsed":false,"visibility":false,"paragraphFormat":{"alignment":"Centered","firstLineIndent":0.0,"keepTogether":false,"keepWithNext":false,"leftIndent":72.0,"lineSpacing":18.0,"lineUnitAfter":0.0,"lineUnitBefore":0.0,"mirrorIndents":false,"outlineLevel":"OutlineLevelBodyText","rightIndent":72.0,"spaceAfter":30.0,"spaceBefore":30.0,"widowControl":true},"font":{"name":"DengXian","size":14.0,"bold":true,"italic":true,"color":"#8DD873","underline":"Single","subscript":false,"superscript":false,"strikeThrough":false,"doubleStrikeThrough":true,"highlightColor":null,"hidden":false},"shading":{"backgroundPatternColor":"#00FF00"}},{"baseStyle":"Table Normal","builtIn":false,"inUse":true,"linked":false,"nextParagraphStyle":"NewTableStyle","nameLocal":"NewTableStyle","priority":100,"type":"Table","unhideWhenUsed":false,"visibility":false,"paragraphFormat":{"alignment":"Left","firstLineIndent":0.0,"keepTogether":false,"keepWithNext":false,"leftIndent":0.0,"lineSpacing":12.0,"lineUnitAfter":0.0,"lineUnitBefore":0.0,"mirrorIndents":false,"outlineLevel":"OutlineLevelBodyText","rightIndent":0.0,"spaceAfter":0.0,"spaceBefore":0.0,"widowControl":true},"font":{"name":"DengXian","size":20.0,"bold":false,"italic":true,"color":"#D86DCB","underline":"None","subscript":false,"superscript":false,"strikeThrough":false,"doubleStrikeThrough":false,"highlightColor":null,"hidden":false},"tableStyle":{"allowBreakAcrossPage":true,"alignment":"Left","bottomCellMargin":0.0,"leftCellMargin":0.08,"rightCellMargin":0.08,"topCellMargin":0.0,"cellSpacing":0.0},"shading":{"backgroundPatternColor":"#60CAF3"}}]}';
    const styles = context.document.importStylesFromJson(str);
    await context.sync();
    console.log("Styles imported from JSON:", styles);
  });
}

// Default helper for invoking an action and handling errors.
async function tryCatch(callback) {
  try {
    await callback();
  } catch (error) {
    console.error(error);
  }
}
```

