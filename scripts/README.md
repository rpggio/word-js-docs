# Word JS Docs Generator

This directory contains scripts for generating documentation from the Word JS samples.

## Available Scripts

### Generate Markdown Documentation

The `gen-markdown-docs.ts` script generates markdown documentation from YAML sample files in the `samples` directory.

To run:

```bash
npm run gen-docs
```

This will:

1. Read all `.yaml` files from the `samples` directory
2. Extract the name, description, template, and script sections
3. Convert HTML in the template section to regular markdown text
4. Strip global-scope jQuery calls from script code
5. Generate markdown documentation in the `dist` directory, preserving the original folder structure

### Generated Documentation Structure

Each markdown file will include:

- Name (as a title)
- Description
- Template content (converted from HTML to markdown)
- Script code (with proper syntax highlighting and jQuery calls removed)

The output directory structure in `dist` will mirror the input structure in `samples`.
