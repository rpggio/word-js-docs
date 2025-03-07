# Insert headers and footers

Inserts headers and footers in the document.

This sample inserts headers and footers in the document.

```typescript
async function addHeader() {
  await Word.run(async (context) => {
    context.document.sections
      .getFirst()
      .getHeader(Word.HeaderFooterType.primary)
      .insertParagraph("This is a primary header.", "End");

    await context.sync();
  });
}

async function addFooter() {
  await Word.run(async (context) => {
    context.document.sections
      .getFirst()
      .getFooter("Primary")
      .insertParagraph("This is a primary footer.", "End");

    await context.sync();
  });
}

async function addFirstPageHeader() {
  await Word.run(async (context) => {
    context.document.sections
      .getFirst()
      .getHeader("FirstPage")
      .insertParagraph("This is a first-page header.", "End");

    await context.sync();
  });
}

async function addFirstPageFooter() {
  await Word.run(async (context) => {
    context.document.sections
      .getFirst()
      .getFooter(Word.HeaderFooterType.firstPage)
      .insertParagraph("This is a first-page footer.", "End");

    await context.sync();
  });
}

async function addEvenPagesHeader() {
  await Word.run(async (context) => {
    context.document.sections
      .getFirst()
      .getHeader(Word.HeaderFooterType.evenPages)
      .insertParagraph("This is an even-pages header.", "End");

    await context.sync();
  });
}

async function addEvenPagesFooter() {
  await Word.run(async (context) => {
    context.document.sections
      .getFirst()
      .getFooter("EvenPages")
      .insertParagraph("This is an even-pages footer.", "End");

    await context.sync();
  });
}

async function setup() {
  await Word.run(async (context) => {
    // Set up text in the document body.
    const body: Word.Body = context.document.body;
    body.clear();
    body.insertParagraph(
      "HeaderFooterType.firstPage applies the header or footer to the first page of the current section. HeaderFooterType.evenPages applies the header or footer to the even pages of the current section. By default, HeaderFooterType.primary applies the header or footer to all pages in the current section. However, if either or both options for FirstPage and EvenPages are set, Primary excludes those pages.",
      "Start"
    );
    body.insertBreak(Word.BreakType.page, Word.InsertLocation.end);
    body.insertParagraph(
      "Do you want to create a solution that extends the functionality of Word? You can use the Office Add-ins platform to extend Word clients running on the web, on a Windows desktop, or on a Mac.",
      "End"
    );
    body.insertBreak(Word.BreakType.page, Word.InsertLocation.end);
    body.insertParagraph(
      "Use add-in commands to extend the Word UI and launch task panes that run JavaScript that interacts with the content in a Word document. Any code that you can run in a browser can run in a Word add-in. Add-ins that interact with content in a Word document create requests to act on Word objects and synchronize object state.",
      "End"
    );

    // Clear any headers and footers.
    const section: Word.Section = context.document.sections.getFirst();

    section.getHeader("Primary").clear();
    section.getHeader("FirstPage").clear();
    section.getHeader("EvenPages").clear();

    section.getFooter("Primary").clear();
    section.getFooter("FirstPage").clear();
    section.getFooter("EvenPages").clear();
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

