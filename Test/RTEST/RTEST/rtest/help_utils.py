import tkinter as tk
from tkinter import ttk
from help_content import HELP_CONTENT

def show_help_window(root, section):
    """Display help information for the specified section"""
    # Get help content for the section
    section_data = HELP_CONTENT.get(section, {})
    title = section_data.get("title", section.replace('_', ' ').title())
    content = section_data.get("content", [])
    
    # Create a new window
    help_window = tk.Toplevel(root)
    help_window.title(f"Help - {title}")
    help_window.geometry("550x550")
    help_window.transient(root)
    help_window.grab_set()  # Make window modal
    
    # Add a frame with scrollable text
    frame = ttk.Frame(help_window, padding="10")
    frame.pack(fill=tk.BOTH, expand=True)
    
    # Scrollable text area
    scrollbar = ttk.Scrollbar(frame)
    scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
    
    text_area = tk.Text(frame, wrap=tk.WORD, yscrollcommand=scrollbar.set)
    text_area.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
    scrollbar.config(command=text_area.yview)
    
    # Configure text tags for formatting
    text_area.tag_configure("heading", font=("", 12, "bold"))
    text_area.tag_configure("subheading", font=("", 11, "bold"))
    text_area.tag_configure("normal", font=("", 10))
    
    # Insert the content with appropriate tags
    for tag, text in content:
        text_area.insert(tk.END, text, tag)
    
    # Make text read-only
    text_area.config(state=tk.DISABLED)
    
    # Add close button
    close_button = ttk.Button(help_window, text="Close", command=help_window.destroy)
    close_button.pack(pady=10)