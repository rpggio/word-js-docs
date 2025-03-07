# Manage settings

This sample shows how to add, edit, get, and delete custom settings on a document.

This sample shows how to add, edit, get, and delete custom settings on a document. Settings created by an add-in
        can
        only be managed by that add-in.

```typescript
async function addEditSetting() {
  // Adds a new custom setting or
  // edits the value of an existing one.
  await Word.run(async (context) => {
    const key = $("#key")
      .val()
      .toString();

    if (key == "") {
      console.error("Key shouldn't be empty.");
      return;
    }

    const value = $("#value")
      .val()
      .toString();

    const settings: Word.SettingCollection = context.document.settings;
    const setting: Word.Setting = settings.add(key, value);
    setting.load();
    await context.sync();

    console.log("Setting added or edited:", setting);
  });
}

async function getAllSettings() {
  // Gets all custom settings this add-in set on this document.
  await Word.run(async (context) => {
    const settings: Word.SettingCollection = context.document.settings;
    settings.load("items");
    await context.sync();

    if (settings.items.length == 0) {
      console.log("There are no settings.");
    } else {
      console.log("All settings:");
      for (let i = 0; i < settings.items.length; i++) {
        console.log(settings.items[i]);
      }
    }
  });
}

async function deleteAllSettings() {
  // Deletes all custom settings this add-in had set on this document.
  await Word.run(async (context) => {
    const settings: Word.SettingCollection = context.document.settings;
    settings.deleteAll();
    await context.sync();
    console.log("All settings deleted.");
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

