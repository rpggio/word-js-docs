#!/bin/bash

# Output file
README="README.md"

# Backup the original README
cp "$README" "${README}.bak"

# Get the content up to the "### Samples" section
sed -n '/^### Samples/q;p' "${README}.bak" > "$README"

# Append the "### Samples" heading
echo "### Samples" >> "$README"
echo "" >> "$README"

# Process each directory in samples (sort them numerically)
find samples -maxdepth 1 -type d | grep -v "^samples$" | sort | while read -r dir; do
    # Extract the directory name without the path
    dirname=$(basename "$dir")
    
    # Extract the category name from the directory name (remove leading numbers and dash)
    category=$(echo "$dirname" | sed 's/^[0-9]*-//')
    
    # Handle category names based on specific categories
    case "$category" in
        "basics")
            category_title="Basics"
            ;;
        "content-controls")
            category_title="Content Controls"
            ;;
        "images")
            category_title="Images"
            ;;
        "lists")
            category_title="Lists"
            ;;
        "paragraph")
            category_title="Paragraph"
            ;;
        "properties")
            category_title="Properties"
            ;;
        "ranges")
            category_title="Ranges"
            ;;
        "tables")
            category_title="Tables"
            ;;
        "document")
            category_title="Document"
            ;;
        "scenarios")
            category_title="Scenarios"
            ;;
        "preview-apis")
            category_title="Preview APIs"
            ;;
        *)
            # Default case: capitalize first letter
            category_title=$(echo "$category" | sed 's/^\(.\)/\u\1/')
            ;;
    esac
    
    # Add category heading
    echo "#### $category_title" >> "$README"
    echo "" >> "$README"
    
    # Find all yaml files in this directory
    find "$dir" -name "*.yaml" | sort | while read -r sample_file; do
        # Extract sample name and description using grep and sed
        name=$(grep "^name:" "$sample_file" | sed 's/^name: //')
        description=$(grep "^description:" "$sample_file" | sed 's/^description: //')
        id=$(grep "^id:" "$sample_file" | sed 's/^id: //')
        
        # Remove any quotes if present
        name=$(echo "$name" | sed 's/"//g')
        description=$(echo "$description" | sed 's/"//g')
        
        # Create a relative link to the sample file
        link="${sample_file}"
        
        # Add bullet point with link and description
        echo "- [$name]($link) - $description" >> "$README"
    done
    
    # Add a blank line after each category
    echo "" >> "$README"
done

echo "README.md has been updated with all samples!" 