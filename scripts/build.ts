import * as fs from "fs-extra";
import * as path from "path";
import * as yaml from "js-yaml";
import * as glob from "glob";

// Function to convert HTML template to markdown
function htmlToMarkdown(html: string): string {
  // Remove setup and samples sections
  html = html.replace(
    /<section[^>]*class="[^"]*setup[^"]*"[^>]*>[\s\S]*?<\/section>/gi,
    ""
  );
  html = html.replace(
    /<section[^>]*class="[^"]*samples[^"]*"[^>]*>[\s\S]*?<\/section>/gi,
    ""
  );

  let markdown = html;

  // Convert common HTML elements to markdown
  markdown = markdown
    // Convert section tags
    .replace(/<\/?section[^>]*>/gi, "")
    // Convert buttons
    .replace(/<button[^>]*>([\s\S]*?)<\/button>/gi, (_, content) => {
      // Extract button label if it exists
      const labelMatch = content.match(/<span[^>]*>(.*?)<\/span>/i);
      return labelMatch ? `**Button:** ${labelMatch[1].trim()}` : "";
    })
    // Convert spans
    .replace(/<span[^>]*>(.*?)<\/span>/gi, "$1")
    // Convert paragraphs
    .replace(/<p[^>]*>([\s\S]*?)<\/p>/gi, "$1\n\n")
    // Convert divs
    .replace(/<div[^>]*>([\s\S]*?)<\/div>/gi, "$1\n")
    // Convert breaks
    .replace(/<br\s*\/?>/gi, "\n")
    // Convert headings
    .replace(/<h1[^>]*>([\s\S]*?)<\/h1>/gi, "# $1\n")
    // Filter out specific h2 headers but keep their content
    .replace(/<h2[^>]*>Description<\/h2>/gi, "")
    .replace(/<h2[^>]*>Template<\/h2>/gi, "")
    .replace(/<h2[^>]*>Script<\/h2>/gi, "")
    // Keep other h2 tags as they were
    .replace(/<h2[^>]*>([\s\S]*?)<\/h2>/gi, "## $1\n")
    .replace(/<h3[^>]*>([\s\S]*?)<\/h3>/gi, "### $1\n")
    .replace(/<h4[^>]*>([\s\S]*?)<\/h4>/gi, "#### $1\n")
    // Convert lists
    .replace(/<ul[^>]*>([\s\S]*?)<\/ul>/gi, "$1\n")
    .replace(/<li[^>]*>([\s\S]*?)<\/li>/gi, "- $1\n")
    // Convert simple formatting
    .replace(/<strong[^>]*>([\s\S]*?)<\/strong>/gi, "**$1**")
    .replace(/<b[^>]*>([\s\S]*?)<\/b>/gi, "**$1**")
    .replace(/<em[^>]*>([\s\S]*?)<\/em>/gi, "*$1*")
    .replace(/<i[^>]*>([\s\S]*?)<\/i>/gi, "*$1*")
    // Remove any remaining HTML tags
    .replace(/<[^>]+>/g, "");

  // Clean up extra whitespace and line breaks
  markdown = markdown.replace(/\n\s*\n\s*\n/g, "\n\n").trim();

  return markdown;
}

// Function to strip global-scope jQuery calls
function stripJQueryCalls(scriptContent: string): string {
  // Remove lines that start with jQuery calls
  return scriptContent
    .split("\n")
    .filter(
      (line) =>
        !line.trim().startsWith("$") && !line.trim().startsWith("jQuery")
    )
    .join("\n")
    .trim();
}

// Function to generate index.md with links to all generated markdown files
async function generateIndexFile(
  markdownFiles: Array<{ path: string; name: string; description: string }>
): Promise<void> {
  // Sort files by their path to group them by section
  markdownFiles.sort((a, b) => a.path.localeCompare(b.path));

  let indexContent = "# Word JS API Docs\n\n";

  // Add link to TypeScript definitions file
  indexContent += "## TypeScript Definitions\n\n";
  indexContent +=
    "- [Office JS TypeScript Definitions](office-js-types.d.ts) - TypeScript type definitions for Office JS APIs\n\n";

  let currentSection = "";

  for (const file of markdownFiles) {
    // Extract section from path (e.g., "50-document" from "docs/50-document/file.md")
    const pathParts = file.path.split("/");
    const section = pathParts.length > 1 ? pathParts[0] : "Other";

    // Add section header if we're in a new section
    if (section !== currentSection) {
      // Format the section name for display (remove numbers, replace hyphens with spaces, capitalize)
      const sectionDisplay = section
        .replace(/^\d+-/, "")
        .replace(/-/g, " ")
        .replace(/\b\w/g, (c) => c.toUpperCase());

      indexContent += `\n## ${sectionDisplay}\n\n`;
      currentSection = section;
    }

    // Add file link and description
    const relativePath = file.path;
    indexContent += `- [${file.name}](${relativePath}) - ${
      file.description || ""
    }\n`;
  }

  // Write the index file
  await fs.writeFile("docs/index.md", indexContent, "utf8");
  console.log("Generated index file: docs/index.md");
}

// Function to convert YAML file to markdown
async function convertYamlToMarkdown(
  yamlFilePath: string,
  outputDir: string
): Promise<void> {
  try {
    // Read YAML file
    const fileContent = await fs.readFile(yamlFilePath, "utf8");
    const data = yaml.load(fileContent) as any;

    // Extract required fields
    const { name, description, template, script } = data;

    if (!name) {
      console.warn(`Skipping ${yamlFilePath}: Missing name property`);
      return;
    }

    // Create markdown content
    let markdownContent = `# ${name}\n\n`;

    if (description) {
      markdownContent += `${description}\n\n`;
    }

    // Convert HTML template to markdown if present
    if (template && template.content) {
      const templateMarkdown = htmlToMarkdown(template.content);
      if (templateMarkdown.trim()) {
        markdownContent += `${templateMarkdown}\n\n`;
      }
    }

    // Process script content
    if (script && script.content) {
      const language = script.language || "javascript";
      const cleanedScript = stripJQueryCalls(script.content);
      markdownContent += `\`\`\`${language}\n${cleanedScript}\n\`\`\`\n\n`;
    }

    // Determine output path based on input path
    const relativePath = path.relative("samples", yamlFilePath);
    const dirName = path.dirname(relativePath);
    const baseName = path.basename(yamlFilePath, ".yaml");

    // Create output directory
    const outputPath = path.join(outputDir, dirName);
    await fs.ensureDir(outputPath);

    // Write markdown file
    const markdownFilePath = path.join(outputPath, `${baseName}.md`);
    await fs.writeFile(markdownFilePath, markdownContent, "utf8");

    console.log(`Converted: ${yamlFilePath} => ${markdownFilePath}`);
  } catch (error) {
    console.error(`Error processing ${yamlFilePath}:`, error);
  }
}

// Main function to process all YAML files
async function main(): Promise<void> {
  try {
    const yamlFiles = glob.sync("samples/**/*.yaml");

    // Clear and recreate the docs directory
    await fs.remove("docs");
    await fs.ensureDir("docs");

    // Copy office-js-types.d.ts to docs directory
    await fs.copy("office-js-types.d.ts", "docs/office-js-types.d.ts");
    console.log("Copied office-js-types.d.ts to docs directory");

    // Store information about each markdown file for the index
    const markdownFiles: Array<{
      path: string;
      name: string;
      description: string;
    }> = [];

    // Process each YAML file
    for (const yamlFile of yamlFiles) {
      // Read the YAML file to extract name and description
      const fileContent = await fs.readFile(yamlFile, "utf8");
      const data = yaml.load(fileContent) as any;

      // Skip files without a name
      if (!data.name) {
        console.warn(`Skipping ${yamlFile}: Missing name property`);
        continue;
      }

      // Determine output path based on input path
      const relativePath = path.relative("samples", yamlFile);
      const dirName = path.dirname(relativePath);
      const baseName = path.basename(yamlFile, ".yaml");
      const markdownPath = path.join(dirName, `${baseName}.md`);

      // Store information for the index
      markdownFiles.push({
        path: markdownPath,
        name: data.name,
        description: data.description || "",
      });

      // Convert YAML to markdown
      await convertYamlToMarkdown(yamlFile, "docs");
    }

    // Generate index.md file
    await generateIndexFile(markdownFiles);

    console.log("Markdown generation complete");
  } catch (error) {
    console.error("Error generating markdown docs:", error);
    process.exit(1);
  }
}

// Run the script
main();
