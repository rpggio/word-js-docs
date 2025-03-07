# Get text

Shows how to get paragraph text, including hidden text and text marked for deletion.

This sample demonstrates how to get paragraph text, including hidden text and text marked for deletion.

        - How to hide selected text (only available on Windows):
            
                Open the **Font** dialog (e.g., right-click the text then select Font from the context menu).

                - Turn on the **Hidden** checkbox.

                - Choose **OK**.

            - How to Track changes in
                Word.

```typescript
async function run() {
  await Word.run(async (context) => {
    const paragraph: Word.Paragraph = context.document.getSelection().paragraphs.getFirst();

    const text = paragraph.getText();
    const textIncludingHidden = paragraph.getText({ IncludeHiddenText: true });
    const textIncludingDeleted = paragraph.getText({ IncludeTextMarkedAsDeleted: true });

    await context.sync();

    console.log("Text:- " + text.value, "Including hidden text:- " + textIncludingHidden.value, "Including text marked as deleted:- " + textIncludingDeleted.value);
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
    body.insertParagraph(
      "Use add-in commands to extend the Word UI and launch task panes that run JavaScript that interacts with the content in a Word document. Any code that you can run in a browser can run in a Word add-in. Add-ins that interact with content in a Word document create requests to act on Word objects and synchronize object state.",
      "End"
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

