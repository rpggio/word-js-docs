# Word JS API Docs

## TypeScript Definitions

- [Office JS TypeScript Definitions](office-js-types.d.ts) - TypeScript type definitions for Office JS APIs


## Basics

- [Basic API call (JavaScript)](01-basics/basic-api-call-es5.md) - Performs a basic Word API call using plain JavaScript & Promises.
- [Basic API call (TypeScript)](01-basics/basic-api-call.md) - Performs a basic Word API call using TypeScript.
- [Basic API call (Office 2013)](01-basics/basic-common-api-call.md) - Performs a basic Word API call using JavaScript with the "common API" syntax (compatible with Office 2013).

## Content Controls

- [On adding content controls](10-content-controls/content-control-onadded-event.md) - Registers, triggers, and deregisters onAdded event that tracks the addition of content controls.
- [On changing data in content controls](10-content-controls/content-control-ondatachanged-event.md) - Registers, triggers, and deregisters onDataChanged event that tracks when data is changed in content controls.
- [On deleting content controls](10-content-controls/content-control-ondeleted-event.md) - Registers, triggers, and deregisters onDeleted event that tracks the removal of content controls.
- [On entering content controls](10-content-controls/content-control-onentered-event.md) - Registers, triggers, and deregisters onEntered event that tracks when the cursor is placed within content controls.
- [On exiting content controls](10-content-controls/content-control-onexited-event.md) - Registers, triggers, and deregisters onExited event that tracks when the cursor is removed from within content controls.
- [On changing selection in content controls](10-content-controls/content-control-onselectionchanged-event.md) - Registers, triggers, and deregisters onSelectionChanged event that tracks when selections are changed in content controls.
- [Get change tracking states of content controls](10-content-controls/get-change-tracking-states.md) - Gets change tracking states of content controls.
- [Manage checkbox content controls](10-content-controls/insert-and-change-checkbox-content-control.md) - Inserts, updates, retrieves, and deletes checkbox content controls.
- [Manage combo box content controls](10-content-controls/insert-and-change-combo-box-content-control.md) - Inserts, updates, and deletes combo box content controls.
- [Content control basics](10-content-controls/insert-and-change-content-controls.md) - Inserts, updates, and retrieves content controls.
- [Manage dropdown list content controls](10-content-controls/insert-and-change-dropdown-list-content-control.md) - Inserts, updates, and deletes dropdown list content controls.

## Images

- [Use inline pictures](15-images/insert-and-get-pictures.md) - Inserts and gets inline pictures.

## Lists

- [Create a list](20-lists/insert-list.md) - Inserts a new list into the document.
- [Get list styles](20-lists/manage-list-styles.md) - This sample shows how to get the list styles in the current document.
- [Organize a list](20-lists/organize-list.md) - Shows how to create and organize a list.

## Paragraph

- [Get paragraph from insertion point](25-paragraph/get-paragraph-on-insertion-point.md) - Gets the full paragraph containing the insertion point.
- [Get text](25-paragraph/get-text.md) - Shows how to get paragraph text, including hidden text and text marked for deletion.
- [Get word count](25-paragraph/get-word-count.md) - Counts how many times a word or term appears in the document.
- [Insert formatted text](25-paragraph/insert-formatted-text.md) - Formats text with pre-built and custom styles.
- [Insert headers and footers](25-paragraph/insert-header-and-footer.md) - Inserts headers and footers in the document.
- [Insert content at different locations](25-paragraph/insert-in-different-locations.md) - Inserts content at different document locations.
- [Insert breaks](25-paragraph/insert-line-and-page-breaks.md) - Inserts page and line breaks in a document.
- [On adding paragraphs](25-paragraph/onadded-event.md) - Registers, triggers, and deregisters the onParagraphAdded event that tracks the addition of paragraphs.
- [On changing content in paragraphs](25-paragraph/onchanged-event.md) - Registers, triggers, and deregisters the onParagraphChanged event that tracks when content is changed in paragraphs.
- [On deleting paragraphs](25-paragraph/ondeleted-event.md) - Registers, triggers, and deregisters the onParagraphDeleted event that tracks the removal of paragraphs.
- [Paragraph properties](25-paragraph/paragraph-properties.md) - Sets indentation, space between paragraphs, and other paragraph properties.
- [Search](25-paragraph/search.md) - Shows basic and advanced search capabilities.

## Properties

- [Built-in document properties](30-properties/get-built-in-properties.md) - Gets built-in document properties.
- [Custom document properties](30-properties/read-write-custom-document-properties.md) - Adds and reads custom document properties of different types.

## Ranges

- [Compare range locations](35-ranges/compare-location.md) - This sample shows how to compare the locations of two ranges.
- [Scroll to a range](35-ranges/scroll-to-range.md) - Scrolls to a range with and without selection.
- [Split a paragraph into ranges](35-ranges/split-words-of-first-paragraph.md) - Splits a paragraph into word ranges and then traverses all the ranges to format each word, producing a "karaoke" effect.

## Tables

- [Manage custom table style](40-tables/manage-custom-style.md) - Shows how to manage primarily margins and alignments of a custom table style in the current document.
- [Table formatting](40-tables/manage-formatting.md) - Gets the formatting details of a table, a table row, and a table cell, including borders, alignment, and cell padding.
- [Create and access a table](40-tables/table-cell-access.md) - Creates a table and accesses a specific cell.

## Document

- [Compare documents](50-document/compare-documents.md) - Compares two documents (the current one and a specified external one).
- [Get styles from external document](50-document/get-external-styles.md) - This sample shows how to get styles from an external document.
- [Insert an external document](50-document/insert-external-document.md) - Inserts the content (with or without settings) of an external document into the current document. Settings include formatting, change-tracking mode, custom properties, and XML parts.
- [Add a section](50-document/insert-section-breaks.md) - Shows how to insert section breaks in the document.
- [Manage annotations](50-document/manage-annotations.md) - Shows how to leverage the Writing Assistance API to manage annotations and use annotation events.
- [Manage body](50-document/manage-body.md) - Shows how to manage the document body.
- [Track changes](50-document/manage-change-tracking.md) - This sample shows how to get and set the change tracking mode and get the before and after of reviewed text.
- [Manage comments](50-document/manage-comments.md) - This sample shows how to perform basic comments operations, including insert, reply, get, edit, resolve, and delete.
- [Manage a CustomXmlPart with the namespace](50-document/manage-custom-xml-part-ns.md) - This sample shows how to add, query, replace, edit, and delete a custom XML part in a document.
- [Manage a CustomXmlPart without the namespace](50-document/manage-custom-xml-part.md) - This sample shows how to add, query, edit, and delete a custom XML part in a document.
- [Manage fields](50-document/manage-fields.md) - This sample shows how to perform basic operations on fields, including insert, get, and delete.
- [Manage footnotes](50-document/manage-footnotes.md) - This sample shows how to perform basic footnote operations, including insert, get, and delete.
- [Manage settings](50-document/manage-settings.md) - This sample shows how to add, edit, get, and delete custom settings on a document.
- [Manage styles](50-document/manage-styles.md) - This sample shows how to perform operations on the styles in the current document and how to add and delete custom styles.
- [Manage tracked changes](50-document/manage-tracked-changes.md) - This samples shows how to manage tracked changes, including accepting and rejecting changes.
- [Manage document save and close](50-document/save-close.md) - Shows how to manage saving and closing document.

## Scenarios

- [Correlated objects pattern](90-scenarios/correlated-objects-pattern.md) - Shows the performance benefits of avoiding `context.sync` calls in a loop.
- [Document assembly](90-scenarios/doc-assembly.md) - Composes different parts of a Word document.
- [Set multiple properties at once](90-scenarios/multiple-property-set.md) - Sets multiple properties at once with the API object set() method.

## Preview Apis

- [Content control basics](99-preview-apis/insert-and-change-content-controls.md) - Inserts, updates, and retrieves content controls.
- [Manage comments](99-preview-apis/manage-comments.md) - This sample shows how to perform operations on comments (including insert, reply, get, edit, resolve, and delete) and use comment events.

## Other

- [Blank snippet](default.md) - Creates a new snippet from a blank template.
