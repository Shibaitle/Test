import tkinter as tk
from tkinter import ttk

def show_help_window(parent, section):
    """Display help information for the specified section"""
    
    # Define help content for each section
    help_content = {
        "text_comparison": {
            "title": "Text File Comparison Tool",
            "content": """
🚀 TEXT FILE COMPARISON TOOL

This powerful tool helps you compare any text-based files with professional-grade features.

📋 SUPPORTED FILE TYPES:
• Text files (.txt, .log)
• Data files (.csv, .json, .xml)
• Code files (.py, .js, .html, .css)
• Configuration files (.ini, .cfg, .md)
• Any text-based file (.* for all files)

✨ KEY FEATURES:
• Side-by-side comparison with color highlighting
• Unified diff view (like Git diff)
• Professional HTML reports
• Advanced comparison options
• Multiple encoding support
• Context line configuration

🎯 HOW TO USE:
1. Select file extension filter (or use .* for all files)
2. Browse and select your original file
3. Browse and select the new/modified file
4. Configure comparison options as needed
5. Click "🔍 Compare Files" to analyze differences
6. Review results in multiple view formats

📊 COMPARISON OPTIONS:
• Ignore whitespace differences
• Case-insensitive comparison  
• Show/hide line numbers
• Adjustable context lines
• Save results to file

💾 OUTPUT FORMATS:
• Summary view with statistics
• Unified diff (text format)
• Side-by-side visual comparison
• HTML reports for sharing

🔧 ADVANCED FEATURES:
• Automatic encoding detection
• Large file support
• Search within results
• Export capabilities
• Professional formatting

This tool seamlessly integrates with the Excel comparison tool, providing
a complete solution for all your file comparison needs.
"""
        },
        # ... existing help content for other sections ...
    }
    
    # Add the existing help content here
    help_content.update({
        "mode_selection": {
            "title": "Comparison Mode Selection",
            "content": """
Choose between two different comparison modes:

STANDARD MODE:
• Uses predefined columns: Team, App Name, and Category
• Best for standardized Excel templates
• Includes built-in filter options
• Supports additional criteria columns

CUSTOM MODE:
• Define your own key columns for matching rows
• Flexible for any Excel structure
• Custom filter criteria based on your key columns
• More complex but adaptable to any format

Select the mode that best fits your Excel file structure.
"""
        },
        
        "file_selection": {
            "title": "File Selection Help",
            "content": """
SELECT YOUR FILES:

Old File (to update):
• This is the Excel file that will be modified
• Choose the file you want to update with new data
• Make sure the file is not open in Excel

New File (reference):
• This is the Excel file with the updated/new data
• The tool will copy data from this file to the old file
• This file remains unchanged

IMPORTANT NOTES:
• Both files must have the same sheet structure
• Column headers should match between files
• Files should use the same data format
• Close both files in Excel before comparing
"""
        },
        
        "sheet_selection": {
            "title": "Sheet Selection Help",
            "content": """
SHEET SELECTION:

1. Click "Load Sheets" to find common sheets between both files
2. Select one or more sheets to process by checking the boxes
3. Click "Load Columns" to populate column dropdowns for selected sheets

NOTES:
• Only sheets that exist in both files will be shown
• You can process multiple sheets in one operation
• Each sheet will be processed with the same criteria and filters
• The header row setting applies to all selected sheets
"""
        },
        
        "comparison_criteria": {
            "title": "Comparison Criteria Help",
            "content": """
COMPARISON CRITERIA:

Header Row:
• Specify which row contains column headers (default: 4)
• This helps the tool identify column names correctly

Column Selection:
• Team Column: Select the column containing team information
• App Name Column: Select the column with application names  
• Category Column: Select the column with category data

Additional Criteria:
• Click "+ Add Criteria Column" to add more matching columns
• Useful for complex matching requirements
• Each additional column becomes part of the row matching key

Formula-Aware Processing:
• When enabled, the tool detects and preserves formula columns
• Prevents overwriting calculated values
• Updates source data that formulas reference instead

MATCHING PROCESS:
Rows are matched based on the combination of all selected criteria columns.
Only rows with matching criteria will be updated.
"""
        },
        
        "filter_criteria": {
            "title": "Filter Criteria Help",
            "content": """
FILTER CRITERIA:

Use filters to limit which rows are processed:

Basic Filters:
• Team Filters: Only process rows with specific team values
• App Name Filters: Only process rows with specific app names
• Category Filters: Only process rows with specific categories

Adding Filters:
• Click the "+" button to add multiple filter values for each type
• Click the "🔍" button to see available values from your file
• Leave filters empty to process all rows

Additional Filter Criteria:
• Click "➕ Add Filter Criteria" to create filters for additional columns
• Useful when you have custom comparison criteria
• Each filter type works independently

FILTER LOGIC:
• Multiple filters of the same type work as OR (any match passes)
• Different filter types work as AND (all must match)
• Empty filters are ignored (no filtering applied)

Example: If you set Team filter to "Development" and App filter to "WebApp", 
only rows with Team="Development" AND App="WebApp" will be processed.
"""
        },
        
        "save_options": {
            "title": "Save Options Help",
            "content": """
SAVE OPTIONS:

File Creation Mode:
• Create new updated file: Saves results to a new file (recommended)
• Replace original file: Overwrites the original file (use with caution)

Additional Options:
• Create highlighted file: Generates an Excel file with changes highlighted in yellow
• Show update popup: Displays which rows were updated after processing
• Clear file selection: Automatically clears file paths after successful update

OUTPUT FILES:
When creating new files, the tool automatically generates descriptive names:
• Includes filter information if filters were used
• Indicates the comparison mode used
• Preserves the original file extension

HIGHLIGHTED FILES:
• Changes are highlighted in yellow
• Comments show previous values
• Useful for reviewing what changed
• Does not affect the main output file

The tool preserves all Excel formatting, formulas, and comments during updates.
"""
        },
        
        "custom_comparison_criteria": {
            "title": "Custom Comparison Criteria Help",
            "content": """
CUSTOM COMPARISON CRITERIA:

This mode allows flexible row matching using any columns as keys.

Key Columns:
• Define which columns uniquely identify each row
• Rows are matched when ALL key column values are identical
• You can use multiple key columns for complex matching

Steps:
1. Set the header row number
2. Click "Load Columns" to load available columns
3. Click "+ Add Key Column" to add matching criteria
4. Select the appropriate column for each key

Examples:
• Single key: Use "ID" column for unique identifier matching
• Multiple keys: Use "Department" + "Employee Name" for composite matching
• Complex keys: Use "Project" + "Phase" + "Task" for detailed matching

IMPORTANT:
• At least one key column is required
• All key columns must exist in both files
• Key combinations should be unique within each file
• Missing or empty key values are ignored

This mode is ideal for files with non-standard structures or when you need 
more control over how rows are matched between files.
"""
        },
        
        "custom_filter_criteria": {
            "title": "Custom Filter Criteria Help",
            "content": """
CUSTOM FILTER CRITERIA:

Create filters based on your custom key columns.

Adding Custom Filters:
1. Ensure you have defined key columns first
2. Click "+ Add Custom Filter"
3. Select which key column to filter on
4. Add filter values for that column

Filter Management:
• Each key column can have its own filter
• Add multiple values per filter using the "+" button
• Use "🔍" to browse available values from your file
• Remove unwanted filter values with "❌"

Filter Logic:
• Multiple values for the same column work as OR
• Multiple column filters work as AND
• Empty filters are ignored

Example:
If you have key columns "Department" and "Status", you can:
• Filter Department for "Sales" OR "Marketing" 
• AND filter Status for "Active"
• This processes only active sales and marketing records

Custom filters give you precise control over which data gets updated
based on your specific key column structure.
"""
        }
    })
    
    if section not in help_content:
        section = "general"
        help_content["general"] = {
            "title": "General Help",
            "content": "No specific help available for this section."
        }
    
    # Create help window
    help_window = tk.Toplevel(parent)
    help_window.title(f"Help - {help_content[section]['title']}")
    
    # Calculate help window size based on parent window
    try:
        parent.update_idletasks()  # Ensure parent dimensions are current
        parent_width = parent.winfo_width()
        parent_height = parent.winfo_height()
        
        # Set help window to 60% of parent size, with minimums
        help_width = max(600, int(parent_width * 0.6))
        help_height = max(500, int(parent_height * 0.7))
        
        # Center relative to parent
        parent_x = parent.winfo_rootx()
        parent_y = parent.winfo_rooty()
        help_x = parent_x + (parent_width - help_width) // 2
        help_y = parent_y + (parent_height - help_height) // 2
        
        help_window.geometry(f"{help_width}x{help_height}+{help_x}+{help_y}")
        
    except tk.TclError:
        # Fallback to fixed size if parent dimensions unavailable
        help_window.geometry("700x600")
    
    help_window.transient(parent)
    help_window.grab_set()
    help_window.resizable(True, True)
    help_window.minsize(500, 400)
    
    # Create scrollable text widget
    frame = ttk.Frame(help_window)
    frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
    
    text_widget = tk.Text(frame, wrap=tk.WORD, font=("", 10))
    scrollbar = ttk.Scrollbar(frame, orient=tk.VERTICAL, command=text_widget.yview)
    text_widget.configure(yscrollcommand=scrollbar.set)
    
    # Pack widgets
    text_widget.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
    scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
    
    # Insert help content
    text_widget.insert(1.0, help_content[section]['content'])
    text_widget.config(state=tk.DISABLED)
    
    # Add close button
    ttk.Button(help_window, text="Close", command=help_window.destroy).pack(pady=10)
    
    # Focus on help window
    help_window.focus_set()