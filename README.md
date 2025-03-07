# Word JS docs

Extract of the Word-specific TypeScript defs and JS API samples, for your AI to ingest.

Sources:

- https://www.npmjs.com/package/@microsoft/office-js
- https://github.com/OfficeDev/office-js-snippets/tree/main/samples/word

Created 2/7/2025, not synchronized with latest.

## Content

### Types

[Office TypeScript defs](./office-js-types.d.ts)

### Samples

#### Basics

- [Basic API call (JavaScript)](samples/01-basics/basic-api-call-es5.yaml) - Performs a basic Word API call using plain JavaScript & Promises.
- [Basic API call (TypeScript)](samples/01-basics/basic-api-call.yaml) - Performs a basic Word API call using TypeScript.
- [Basic API call (Office 2013)](samples/01-basics/basic-common-api-call.yaml) - Performs a basic Word API call using JavaScript with the common API syntax (compatible with Office 2013).

#### Content Controls

- [On adding content controls](samples/10-content-controls/content-control-onadded-event.yaml) - 'Registers, triggers, and deregisters onAdded event that tracks the addition of content controls.'
- [On changing data in content controls](samples/10-content-controls/content-control-ondatachanged-event.yaml) - 'Registers, triggers, and deregisters onDataChanged event that tracks when data is changed in content controls.'
- [On deleting content controls](samples/10-content-controls/content-control-ondeleted-event.yaml) - 'Registers, triggers, and deregisters onDeleted event that tracks the removal of content controls.'
- [On entering content controls](samples/10-content-controls/content-control-onentered-event.yaml) - 'Registers, triggers, and deregisters onEntered event that tracks when the cursor is placed within content controls.'
- [On exiting content controls](samples/10-content-controls/content-control-onexited-event.yaml) - 'Registers, triggers, and deregisters onExited event that tracks when the cursor is removed from within content controls.'
- [On changing selection in content controls](samples/10-content-controls/content-control-onselectionchanged-event.yaml) - 'Registers, triggers, and deregisters onSelectionChanged event that tracks when selections are changed in content controls.'
- [Get change tracking states of content controls](samples/10-content-controls/get-change-tracking-states.yaml) - Gets change tracking states of content controls.
- [Manage checkbox content controls](samples/10-content-controls/insert-and-change-checkbox-content-control.yaml) - 'Inserts, updates, retrieves, and deletes checkbox content controls.'
- [Manage combo box content controls](samples/10-content-controls/insert-and-change-combo-box-content-control.yaml) - 'Inserts, updates, and deletes combo box content controls.'
- [Content control basics](samples/10-content-controls/insert-and-change-content-controls.yaml) - 'Inserts, updates, and retrieves content controls.'
- [Manage dropdown list content controls](samples/10-content-controls/insert-and-change-dropdown-list-content-control.yaml) - 'Inserts, updates, and deletes dropdown list content controls.'

#### Images

- [Use inline pictures](samples/15-images/insert-and-get-pictures.yaml) - Inserts and gets inline pictures.

#### Lists

- [Create a list](samples/20-lists/insert-list.yaml) - Inserts a new list into the document.
- [Get list styles](samples/20-lists/manage-list-styles.yaml) - This sample shows how to get the list styles in the current document.
- [Organize a list](samples/20-lists/organize-list.yaml) - Shows how to create and organize a list.

#### Paragraph

- [Get paragraph from insertion point](samples/25-paragraph/get-paragraph-on-insertion-point.yaml) - Gets the full paragraph containing the insertion point.
- [Get text](samples/25-paragraph/get-text.yaml) - 'Shows how to get paragraph text, including hidden text and text marked for deletion.'
- [Get word count](samples/25-paragraph/get-word-count.yaml) - Counts how many times a word or term appears in the document.
- [Insert formatted text](samples/25-paragraph/insert-formatted-text.yaml) - Formats text with pre-built and custom styles.
- [Insert headers and footers](samples/25-paragraph/insert-header-and-footer.yaml) - Inserts headers and footers in the document.
- [Insert content at different locations](samples/25-paragraph/insert-in-different-locations.yaml) - Inserts content at different document locations.
- [Insert breaks](samples/25-paragraph/insert-line-and-page-breaks.yaml) - Inserts page and line breaks in a document.
- [On adding paragraphs](samples/25-paragraph/onadded-event.yaml) - 'Registers, triggers, and deregisters the onParagraphAdded event that tracks the addition of paragraphs.'
- [On changing content in paragraphs](samples/25-paragraph/onchanged-event.yaml) - 'Registers, triggers, and deregisters the onParagraphChanged event that tracks when content is changed in paragraphs.'
- [On deleting paragraphs](samples/25-paragraph/ondeleted-event.yaml) - 'Registers, triggers, and deregisters the onParagraphDeleted event that tracks the removal of paragraphs.'
- [Paragraph properties](samples/25-paragraph/paragraph-properties.yaml) - 'Sets indentation, space between paragraphs, and other paragraph properties.'
- [Search](samples/25-paragraph/search.yaml) - Shows basic and advanced search capabilities.

#### Properties

- [Built-in document properties](samples/30-properties/get-built-in-properties.yaml) - Gets built-in document properties.
- [Custom document properties](samples/30-properties/read-write-custom-document-properties.yaml) - Adds and reads custom document properties of different types.

#### Ranges

- [Compare range locations](samples/35-ranges/compare-location.yaml) - This sample shows how to compare the locations of two ranges.
- [Scroll to a range](samples/35-ranges/scroll-to-range.yaml) - Scrolls to a range with and without selection.
- [Split a paragraph into ranges](samples/35-ranges/split-words-of-first-paragraph.yaml) - 'Splits a paragraph into word ranges and then traverses all the ranges to format each word, producing a karaoke effect.'

#### Tables

- [Manage custom table style](samples/40-tables/manage-custom-style.yaml) - Shows how to manage primarily margins and alignments of a custom table style in the current document.
- [Table formatting](samples/40-tables/manage-formatting.yaml) - 'Gets the formatting details of a table, a table row, and a table cell, including borders, alignment, and cell padding.'
- [Create and access a table](samples/40-tables/table-cell-access.yaml) - Creates a table and accesses a specific cell.

#### Document

- [Compare documents](samples/50-document/compare-documents.yaml) - Compares two documents (the current one and a specified external one).
- [Get styles from external document](samples/50-document/get-external-styles.yaml) - This sample shows how to get styles from an external document.
- [Insert an external document](samples/50-document/insert-external-document.yaml) - 'Inserts the content (with or without settings) of an external document into the current document. Settings include formatting, change-tracking mode, custom properties, and XML parts.'
- [Add a section](samples/50-document/insert-section-breaks.yaml) - Shows how to insert section breaks in the document.
- [Manage annotations](samples/50-document/manage-annotations.yaml) - Shows how to leverage the Writing Assistance API to manage annotations and use annotation events.
- [Manage body](samples/50-document/manage-body.yaml) - Shows how to manage the document body.
- [Track changes](samples/50-document/manage-change-tracking.yaml) - This sample shows how to get and set the change tracking mode and get the before and after of reviewed text.
- [Manage comments](samples/50-document/manage-comments.yaml) - 'This sample shows how to perform basic comments operations, including insert, reply, get, edit, resolve, and delete.'
- [Manage a CustomXmlPart with the namespace](samples/50-document/manage-custom-xml-part-ns.yaml) - 'This sample shows how to add, query, replace, edit, and delete a custom XML part in a document.'
- [Manage a CustomXmlPart without the namespace](samples/50-document/manage-custom-xml-part.yaml) - 'This sample shows how to add, query, edit, and delete a custom XML part in a document.'
- [Manage fields](samples/50-document/manage-fields.yaml) - 'This sample shows how to perform basic operations on fields, including insert, get, and delete.'
- [Manage footnotes](samples/50-document/manage-footnotes.yaml) - 'This sample shows how to perform basic footnote operations, including insert, get, and delete.'
- [Manage settings](samples/50-document/manage-settings.yaml) - 'This sample shows how to add, edit, get, and delete custom settings on a document.'
- [Manage styles](samples/50-document/manage-styles.yaml) - This sample shows how to perform operations on the styles in the current document and how to add and delete custom styles.
- [Manage tracked changes](samples/50-document/manage-tracked-changes.yaml) - 'This samples shows how to manage tracked changes, including accepting and rejecting changes.'
- [Manage document save and close](samples/50-document/save-close.yaml) - Shows how to manage saving and closing document.

#### Scenarios

- [Correlated objects pattern](samples/90-scenarios/correlated-objects-pattern.yaml) - Shows the performance benefits of avoiding `context.sync` calls in a loop.
- [Document assembly](samples/90-scenarios/doc-assembly.yaml) - Composes different parts of a Word document.
- [Set multiple properties at once](samples/90-scenarios/multiple-property-set.yaml) - Sets multiple properties at once with the API object set() method.

#### Preview APIs

- [Content control basics](samples/99-preview-apis/insert-and-change-content-controls.yaml) - 'Inserts, updates, and retrieves content controls.'
- [Manage comments](samples/99-preview-apis/manage-comments.yaml) - 'This sample shows how to perform operations on comments (including insert, reply, get, edit, resolve, and delete) and use comment events.'

