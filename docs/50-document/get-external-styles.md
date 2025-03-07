# Get styles from external document

This sample shows how to get styles from an external document.

This sample demonstrates how to get styles from an external document.

```typescript
let externalDocument;

async function getExternalStyles() {
  // Gets style info from another document passed in as a Base64-encoded string.
  await Word.run(async (context) => {
    const retrievedStyles = context.application.retrieveStylesFromBase64(externalDocument);
    await context.sync();

    console.log("Styles from the other document:", retrievedStyles.value);
  });
}

function getBase64() {
  // Retrieve the file and set up an HTML FileReader element.
  const myFile = <HTMLInputElement>document.getElementById("file");
  const reader = new FileReader();

  reader.onload = (event) => {
    // Remove the metadata before the Base64-encoded string.
    const startIndex = reader.result.toString().indexOf("base64,");
    externalDocument = reader.result.toString().substr(startIndex + 7);
  };

  // Read the file as a data URL so that we can parse the Base64-encoded string.
  reader.readAsDataURL(myFile.files[0]);
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

