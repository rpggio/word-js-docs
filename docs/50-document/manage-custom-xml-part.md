# Manage a CustomXmlPart without the namespace

This sample shows how to add, query, edit, and delete a custom XML part in a document.

This sample shows how to add, query, edit, and delete a custom XML part in a document.

  **Note**: For your production add-in, make sure to create and host your own XML schema.

```typescript
async function addCustomXmlPart() {
  // Adds a custom XML part.
  await Word.run(async (context) => {
    const originalXml =
      "<Reviewers><Reviewer>Juan</Reviewer><Reviewer>Hong</Reviewer><Reviewer>Sally</Reviewer></Reviewers>";
    const customXmlPart: Word.CustomXmlPart = context.document.customXmlParts.add(originalXml);
    customXmlPart.load("id");
    const xmlBlob = customXmlPart.getXml();

    await context.sync();

    const readableXml = addLineBreaksToXML(xmlBlob.value);
    console.log("Added custom XML part:", readableXml);

    // Store the XML part's ID in a setting so the ID is available to other functions.
    const settings: Word.SettingCollection = context.document.settings;
    settings.add("ContosoReviewXmlPartId", customXmlPart.id);

    await context.sync();
  });
}

async function query() {
  // Original XML: <Reviewers><Reviewer>Juan</Reviewer><Reviewer>Hong</Reviewer><Reviewer>Sally</Reviewer></Reviewers>
  
  // Queries a custom XML part for elements matching the search terms.
  await Word.run(async (context) => {
    const settings: Word.SettingCollection = context.document.settings;
    const xmlPartIDSetting: Word.Setting = settings.getItemOrNullObject("ContosoReviewXmlPartId").load("value");

    await context.sync();

    if (xmlPartIDSetting.value) {
      const customXmlPart: Word.CustomXmlPart = context.document.customXmlParts.getItem(xmlPartIDSetting.value);
      const xpathToQueryFor = "/Reviewers/Reviewer";
      const clientResult = customXmlPart.query(xpathToQueryFor, {
        contoso: "http://schemas.contoso.com/review/1.0"
      });

      await context.sync();

      console.log(`Queried custom XML part for ${xpathToQueryFor} and found ${clientResult.value.length} matches:`);
      for (let i = 0; i < clientResult.value.length; i++) {
        console.log(clientResult.value[i]);
      }
    } else {
      console.warn("Didn't find custom XML part to query.");
    }
  });
}

async function insertAttribute() {
  // Original XML: <Reviewers><Reviewer>Juan</Reviewer><Reviewer>Hong</Reviewer><Reviewer>Sally</Reviewer></Reviewers>
  
  // Inserts an attribute into a custom XML part.
  await Word.run(async (context) => {
    const settings: Word.SettingCollection = context.document.settings;
    const xmlPartIDSetting: Word.Setting = settings.getItemOrNullObject("ContosoReviewXmlPartId").load("value");
    await context.sync();

    if (xmlPartIDSetting.value) {
      const customXmlPart: Word.CustomXmlPart = context.document.customXmlParts.getItem(xmlPartIDSetting.value);

      // The insertAttribute method inserts an attribute with the given name and value into the element identified by the xpath parameter.
      customXmlPart.insertAttribute("/Reviewers", { contoso: "http://schemas.contoso.com/review/1.0" }, "Nation", "US");
      const xmlBlob = customXmlPart.getXml();
      await context.sync();

      const readableXml = addLineBreaksToXML(xmlBlob.value);
      console.log("Successfully inserted attribute:", readableXml);
    } else {
      console.warn("Didn't find custom XML part to insert attribute into.");
    }
  });
}

async function insertElement() {
  // Original XML: <Reviewers><Reviewer>Juan</Reviewer><Reviewer>Hong</Reviewer><Reviewer>Sally</Reviewer></Reviewers>
  
  // Inserts an element into a custom XML part.
  await Word.run(async (context) => {
    const settings: Word.SettingCollection = context.document.settings;
    const xmlPartIDSetting: Word.Setting = settings.getItemOrNullObject("ContosoReviewXmlPartId").load("value");
    await context.sync();

    if (xmlPartIDSetting.value) {
      const customXmlPart: Word.CustomXmlPart = context.document.customXmlParts.getItem(xmlPartIDSetting.value);

      // The insertElement method inserts the given XML under the parent element identified by the xpath parameter at the provided child position index.
      customXmlPart.insertElement(
        "/Reviewers",
        "<Lead>Mark</Lead>",
        { contoso: "http://schemas.contoso.com/review/1.0" },
        0
      );
      const xmlBlob = customXmlPart.getXml();
      await context.sync();

      const readableXml = addLineBreaksToXML(xmlBlob.value);
      console.log("Successfully inserted element:", readableXml);
    } else {
      console.warn("Didn't find custom XML part to insert element into.");
    }
  });
}

async function deleteCustomXmlPart() {
  // Original XML: <Reviewers><Reviewer>Juan</Reviewer><Reviewer>Hong</Reviewer><Reviewer>Sally</Reviewer></Reviewers>
  
  // Deletes a custom XML part.
  await Word.run(async (context) => {
    const settings: Word.SettingCollection = context.document.settings;
    const xmlPartIDSetting: Word.Setting = settings.getItemOrNullObject("ContosoReviewXmlPartId").load("value");
    await context.sync();

    if (xmlPartIDSetting.value) {
      let customXmlPart: Word.CustomXmlPart = context.document.customXmlParts.getItem(xmlPartIDSetting.value);
      const xmlBlob = customXmlPart.getXml();
      customXmlPart.delete();
      customXmlPart = context.document.customXmlParts.getItemOrNullObject(xmlPartIDSetting.value);

      await context.sync();

      if (customXmlPart.isNullObject) {
        console.log(`The XML part with the ID ${xmlPartIDSetting.value} has been deleted.`);

        // Delete the associated setting too.
        xmlPartIDSetting.delete();

        await context.sync();
      } else {
        const readableXml = addLineBreaksToXML(xmlBlob.value);
        console.error(`This is strange. The XML part with the id ${xmlPartIDSetting.value} wasn't deleted:`, readableXml);
      }
    } else {
      console.warn("Didn't find custom XML part to delete.");
    }
  });
}

function addLineBreaksToXML(xmlBlob: string): string {
  const replaceValue = new RegExp(">");
  return xmlBlob.replace(/></g, "> <");
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

