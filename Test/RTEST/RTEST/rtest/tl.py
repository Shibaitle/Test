import tkinter as tk
from tkinter import ttk, messagebox
import os
import sys
from pathlib import Path

class ToolLauncher:
    def __init__(self, root):
        self.root = root
        self.root.title("Comparison Tools Selection")
        
        # Make the window responsive and centered
        self._setup_window()
        
        # Configure modern styles
        self._configure_styles()
        
        # Create the main UI
        self._create_ui()
        
        # Setup window state management
        self._setup_window_state()
    
    def _setup_window(self):
        """Setup window size and position"""
        # Get screen dimensions
        screen_width = self.root.winfo_screenwidth()
        screen_height = self.root.winfo_screenheight()
        
        # Set window size (85% of screen for horizontal layout)
        window_width = int(screen_width * 0.85)
        window_height = int(screen_height * 0.75)
        
        # Ensure minimum size for horizontal layout
        window_width = max(window_width, 1200)
        window_height = max(window_height, 700)
        
        # Center the window
        center_x = int((screen_width - window_width) / 2)
        center_y = int((screen_height - window_height) / 2)
        
        self.root.geometry(f"{window_width}x{window_height}+{center_x}+{center_y}")
        self.root.minsize(1200, 700)
        self.root.resizable(True, True)
        
        # Set window icon and properties
        try:
            self.root.attributes('-alpha', 1.0)
            self.root.lift()
            self.root.focus_force()
        except tk.TclError:
            pass
    
    def _configure_styles(self):
        """Configure modern ttk styles"""
        style = ttk.Style()
        
        # Configure main button style
        style.configure(
            "Tool.TButton",
            font=("Segoe UI", 11, "bold"),
            foreground="#FFFFFF",
            background="#0078D4",
            borderwidth=0,
            focuscolor='none',
            padding=(20, 15)
        )
        
        style.map("Tool.TButton",
            background=[('active', '#106EBE'), ('pressed', '#005A9E')])
        
        # Configure secondary button style
        style.configure(
            "Secondary.TButton",
            font=("Segoe UI", 9),
            foreground="#323130",
            background="#F3F2F1",
            borderwidth=1,
            focuscolor='none',
            padding=(10, 8)
        )
        
        style.map("Secondary.TButton",
            background=[('active', '#EDEBE9'), ('pressed', '#E1DFDD')])
        
        # Configure card frame style
        style.configure(
            "Card.TFrame",
            background="#FFFFFF",
            borderwidth=2,
            relief="solid"
        )
        
        # Configure title label style
        style.configure(
            "Title.TLabel",
            font=("Segoe UI", 16, "bold"),
            foreground="#323130",
            background="#FAFAFA"
        )
        
        # Configure subtitle label style
        style.configure(
            "Subtitle.TLabel",
            font=("Segoe UI", 10),
            foreground="#605E5C",
            background="#FAFAFA"
        )
    
    def _create_ui(self):
        """Create the main user interface"""
        # Main container with background
        main_container = tk.Frame(self.root, bg="#FAFAFA")
        main_container.pack(fill=tk.BOTH, expand=True)
        
        # Create header
        self._create_header(main_container)
        
        # Create scrollable content area
        self._create_content_area(main_container)
        
        # Create footer
        self._create_footer(main_container)
    
    def _create_header(self, parent):
        """Create the header section"""
        header_frame = tk.Frame(parent, bg="#9607C1", height=100)
        header_frame.pack(fill=tk.X)
        header_frame.pack_propagate(False)
        
        # Header content
        header_content = tk.Frame(header_frame, bg="#9607C1")
        header_content.pack(expand=True, pady=20)
        
        # Main title
        title_label = tk.Label(
            header_content,
            text="Comparison Tools Selection",
            font=("Segoe UI", 24, "bold"),
            foreground="#FFFFFF",
            background="#9607C1"
        )
        title_label.pack()
        
    def _create_content_area(self, parent):
        """Create the main content area with horizontal tool cards and pagination"""
        # Content container
        content_container = tk.Frame(parent, bg="#FAFAFA")
        content_container.pack(fill=tk.BOTH, expand=True, padx=20, pady=20)
        
        # Pagination controls at top
        self._create_pagination_controls(content_container)
        
        # Tools display area
        self.tools_display_frame = tk.Frame(content_container, bg="#FAFAFA")
        self.tools_display_frame.pack(fill=tk.BOTH, expand=True, pady=(10, 0))
        
        # Initialize pagination
        self.current_page = 0
        self.tools_per_page = 3  # Show 3 tools per page
        
        # Create tool cards
        self._create_tool_cards_horizontal()
    
    def _create_pagination_controls(self, parent):
        """Create pagination controls"""
        pagination_frame = tk.Frame(parent, bg="#FAFAFA")
        pagination_frame.pack(fill=tk.X, pady=(0, 10))
        
        # Page info and controls container
        controls_container = tk.Frame(pagination_frame, bg="#FAFAFA")
        controls_container.pack()
        
        # Previous button
        self.prev_btn = tk.Button(
            controls_container,
            text="◀ Previous",
            font=("Segoe UI", 10),
            foreground="#323130",
            background="#F3F2F1",
            borderwidth=1,
            padx=15,
            pady=5,
            command=self._previous_page,
            state=tk.DISABLED
        )
        self.prev_btn.pack(side=tk.LEFT, padx=5)
        
        # Page indicator
        self.page_label = tk.Label(
            controls_container,
            text="Page 1 of 1",
            font=("Segoe UI", 10, "bold"),
            foreground="#323130",
            background="#FAFAFA"
        )
        self.page_label.pack(side=tk.LEFT, padx=20)
        
        # Next button
        self.next_btn = tk.Button(
            controls_container,
            text="Next ▶",
            font=("Segoe UI", 10),
            foreground="#323130",
            background="#F3F2F1",
            borderwidth=1,
            padx=15,
            pady=5,
            command=self._next_page,
            state=tk.DISABLED
        )
        self.next_btn.pack(side=tk.LEFT, padx=5)
        
        # Tools per page selector
        tools_per_page_frame = tk.Frame(controls_container, bg="#FAFAFA")
        tools_per_page_frame.pack(side=tk.RIGHT, padx=20)
        
        tk.Label(
            tools_per_page_frame,
            text="Tools per page:",
            font=("Segoe UI", 9),
            foreground="#605E5C",
            background="#FAFAFA"
        ).pack(side=tk.LEFT, padx=(0, 5))
        
        self.tools_per_page_var = tk.StringVar(value="3")
        tools_combo = ttk.Combobox(
            tools_per_page_frame,
            textvariable=self.tools_per_page_var,
            values=["2", "3", "4", "5"],
            width=5,
            state="readonly"
        )
        tools_combo.pack(side=tk.LEFT)
        tools_combo.bind("<<ComboboxSelected>>", self._on_tools_per_page_change)
    
    def _create_tool_cards_horizontal(self):
        """Create tool cards in horizontal layout with pagination"""
        # Clear existing cards
        for widget in self.tools_display_frame.winfo_children():
            widget.destroy()
        
        # Get tools for current page
        tools = self._get_all_tools()
        start_idx = self.current_page * self.tools_per_page
        end_idx = start_idx + self.tools_per_page
        current_tools = tools[start_idx:end_idx]
        
        if not current_tools:
            # No tools to display
            no_tools_label = tk.Label(
                self.tools_display_frame,
                text="No tools available on this page",
                font=("Segoe UI", 14),
                foreground="#605E5C",
                background="#FAFAFA"
            )
            no_tools_label.pack(expand=True)
            return
        
        # Create horizontal container for tools
        tools_container = tk.Frame(self.tools_display_frame, bg="#FAFAFA")
        tools_container.pack(fill=tk.BOTH, expand=True)
        
        # Configure grid weights for equal distribution
        for i in range(len(current_tools)):
            tools_container.grid_columnconfigure(i, weight=1)
        tools_container.grid_rowconfigure(0, weight=1)
        
        # Create tool cards
        for i, tool in enumerate(current_tools):
            self._create_tool_card_horizontal(tools_container, tool, 0, i)
        
        # Update pagination controls
        self._update_pagination_controls(len(tools))
    
    def _get_all_tools(self):
        """Get all available tools (for future expansion)"""
        tools = [
            {
                "title": "📊 Excel Compare & Update",
                "subtitle": "Advanced Excel File Comparison",
                "description": [
                    "• Compare and update Excel files with intelligent matching",
                    "• Standard mode: Team/App Name/Category columns",
                    "• Custom mode: Flexible column matching",
                    "• Formula-aware processing and relationship detection",
                    "• Advanced filtering and bulk operations",
                    "• Automatic backup and change highlighting"
                ],
                "icon": "📊",
                "action": self._launch_excel_tool,
                "color": "#107C10",
                "status": "Available"
            },
            {
                "title": "📄 Text File Comparison",
                "subtitle": "Comprehensive Text File Analysis",
                "description": [
                    "• Side-by-side comparison with syntax highlighting",
                    "• Support for multiple file formats (TXT, LOG, CSV, JSON, XML, Python, etc.)",
                    "• Multiple encoding support (UTF-8, Latin-1, CP1252, etc.)",
                    "• Unified diff view and HTML reports",
                    "• Advanced filtering and search options",
                    "• Export differences to various formats"
                ],
                "icon": "📄",
                "action": self._launch_text_tool,
                "color": "#2196F3",
                "status": "Available"
            },
            {
                "title": "🖼️ Image & Picture Comparison",
                "subtitle": "OCR Text Extraction & Visual Analysis",
                "description": [
                    "• OCR text extraction with multi-language support (English, Thai)",
                    "• Visual difference detection and highlighting",
                    "• Advanced counterfeit detection algorithms",
                    "• Document verification (Banking slips, Receipts, QR payments)",
                    "• QR code scanning and validation",
                    "• AI-generated content detection"
                ],
                "icon": "🖼️",
                "action": self._launch_image_tool,
                "color": "#FF9800",
                "status": "Available"
            },
            # Add more tools here in the future
            # {
            #     "title": "🗄️ Database Comparison",
            #     "subtitle": "Database Schema & Data Analysis",
            #     "description": [
            #         "• Compare database schemas and structures",
            #         "• Data comparison between tables",
            #         "• SQL diff generation",
            #         "• Migration script creation"
            #     ],
            #     "icon": "🗄️",
            #     "action": self._launch_database_tool,
            #     "color": "#9C27B0",
            #     "status": "Coming Soon"
            # },
            # {
            #     "title": "📊 PDF Document Analysis",
            #     "subtitle": "PDF Content Comparison",
            #     "description": [
            #         "• Extract and compare PDF content",
            #         "• Text and image comparison",
            #         "• Metadata analysis",
            #         "• Change tracking in documents"
            #     ],
            #     "icon": "📊",
            #     "action": self._launch_pdf_tool,
            #     "color": "#F44336",
            #     "status": "Coming Soon"
            # }
        ]
        return tools
    
    def _create_tool_card_horizontal(self, parent, tool_info, row, col):
        """Create an individual tool card for horizontal layout"""
        # Card container with consistent sizing
        card_container = tk.Frame(parent, bg="#FAFAFA")
        card_container.grid(row=row, column=col, sticky="nsew", padx=10, pady=10)
        
        # Card frame with fixed dimensions
        card_frame = tk.Frame(
            card_container,
            bg="#FFFFFF",
            relief="solid",
            borderwidth=2,
            width=350,  # Fixed width for consistency
            height=400  # Fixed height for consistency
        )
        card_frame.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        card_frame.pack_propagate(False)  # Maintain fixed size
        
        # Inner frame for content with padding
        inner_frame = tk.Frame(card_frame, bg="#FFFFFF")
        inner_frame.pack(fill=tk.BOTH, expand=True, padx=15, pady=15)
        
        # Add hover effects
        def on_enter(event):
            card_frame.configure(relief="solid", borderwidth=3)
        
        def on_leave(event):
            card_frame.configure(relief="solid", borderwidth=2)
        
        card_frame.bind("<Enter>", on_enter)
        card_frame.bind("<Leave>", on_leave)
        
        # Icon section
        icon_frame = tk.Frame(inner_frame, bg="#FFFFFF")
        icon_frame.pack(pady=(0, 10))
        
        icon_label = tk.Label(
            icon_frame,
            text=tool_info["icon"],
            font=("Segoe UI", 36),
            background="#FFFFFF"
        )
        icon_label.pack()
        
        # Title
        title_label = tk.Label(
            inner_frame,
            text=tool_info["title"],
            font=("Segoe UI", 12, "bold"),
            foreground="#323130",
            background="#FFFFFF",
            wraplength=300,
            justify=tk.CENTER
        )
        title_label.pack(pady=(0, 5))
        
        # Subtitle
        subtitle_label = tk.Label(
            inner_frame,
            text=tool_info["subtitle"],
            font=("Segoe UI", 9),
            foreground="#605E5C",
            background="#FFFFFF",
            wraplength=300,
            justify=tk.CENTER
        )
        subtitle_label.pack(pady=(0, 10))
        
        # Features section (scrollable for long lists)
        features_frame = tk.Frame(inner_frame, bg="#FFFFFF")
        features_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 10))
        
        # Create a canvas for scrollable features if needed
        if len(tool_info["description"]) > 4:
            features_canvas = tk.Canvas(features_frame, bg="#FFFFFF", height=120, highlightthickness=0)
            features_scrollbar = ttk.Scrollbar(features_frame, orient="vertical", command=features_canvas.yview)
            scrollable_features = tk.Frame(features_canvas, bg="#FFFFFF")
            
            scrollable_features.bind(
                "<Configure>",
                lambda e: features_canvas.configure(scrollregion=features_canvas.bbox("all"))
            )
            
            features_canvas.create_window((0, 0), window=scrollable_features, anchor="nw")
            features_canvas.configure(yscrollcommand=features_scrollbar.set)
            
            features_canvas.pack(side="left", fill="both", expand=True)
            features_scrollbar.pack(side="right", fill="y")
            
            features_container = scrollable_features
        else:
            features_container = features_frame
        
        # Add features
        for feature in tool_info["description"][:6]:  # Limit to 6 features for display
            feature_label = tk.Label(
                features_container,
                text=feature,
                font=("Segoe UI", 8),
                foreground="#323130",
                background="#FFFFFF",
                anchor="w",
                justify="left",
                wraplength=280
            )
            feature_label.pack(anchor="w", pady=1)
        
        # Status indicator
        status_frame = tk.Frame(inner_frame, bg="#FFFFFF")
        status_frame.pack(pady=(5, 10))
        
        status_color = "#107C10" if tool_info['status'] == "Available" else "#D13438"
        status_label = tk.Label(
            status_frame,
            text=f"● {tool_info['status']}",
            font=("Segoe UI", 9, "bold"),
            foreground=status_color,
            background="#FFFFFF"
        )
        status_label.pack()
        
        # Launch button
        button_state = tk.NORMAL if tool_info['status'] == "Available" else tk.DISABLED
        launch_btn = tk.Button(
            inner_frame,
            text="🚀 Launch",
            font=("Segoe UI", 10, "bold"),
            foreground="#FFFFFF",
            background=tool_info["color"] if tool_info['status'] == "Available" else "#CCCCCC",
            borderwidth=0,
            padx=20,
            pady=8,
            cursor="hand2" if tool_info['status'] == "Available" else "arrow",
            command=tool_info["action"] if tool_info['status'] == "Available" else None,
            state=button_state
        )
        launch_btn.pack(side=tk.BOTTOM)
        
        # Add button hover effects only if available
        if tool_info['status'] == "Available":
            def btn_on_enter(event):
                launch_btn.configure(background=self._darken_color(tool_info["color"]))
            
            def btn_on_leave(event):
                launch_btn.configure(background=tool_info["color"])
            
            launch_btn.bind("<Enter>", btn_on_enter)
            launch_btn.bind("<Leave>", btn_on_leave)
    
    def _update_pagination_controls(self, total_tools):
        """Update pagination controls based on current state"""
        total_pages = (total_tools - 1) // self.tools_per_page + 1 if total_tools > 0 else 1
        
        # Update page label
        self.page_label.config(text=f"Page {self.current_page + 1} of {total_pages}")
        
        # Update button states
        self.prev_btn.config(state=tk.NORMAL if self.current_page > 0 else tk.DISABLED)
        self.next_btn.config(state=tk.NORMAL if self.current_page < total_pages - 1 else tk.DISABLED)
    
    def _previous_page(self):
        """Go to previous page"""
        if self.current_page > 0:
            self.current_page -= 1
            self._create_tool_cards_horizontal()
    
    def _next_page(self):
        """Go to next page"""
        tools = self._get_all_tools()
        total_pages = (len(tools) - 1) // self.tools_per_page + 1
        if self.current_page < total_pages - 1:
            self.current_page += 1
            self._create_tool_cards_horizontal()
    
    def _on_tools_per_page_change(self, event=None):
        """Handle change in tools per page"""
        self.tools_per_page = int(self.tools_per_page_var.get())
        self.current_page = 0  # Reset to first page
        self._create_tool_cards_horizontal()

    def _create_footer(self, parent):
        """Create the footer section"""
        footer_frame = tk.Frame(parent, bg="#F3F2F1", height=60)
        footer_frame.pack(fill=tk.X, side=tk.BOTTOM)
        footer_frame.pack_propagate(False)
        
        # Footer content
        footer_content = tk.Frame(footer_frame, bg="#F3F2F1")
        footer_content.pack(expand=True, pady=15)
        
        # Footer buttons
        buttons_frame = tk.Frame(footer_content, bg="#F3F2F1")
        buttons_frame.pack()
        
        # Help button
        help_btn = tk.Button(
            buttons_frame,
            text="❓ Help & Documentation",
            font=("Segoe UI", 9),
            foreground="#323130",
            background="#F3F2F1",
            borderwidth=1,
            padx=15,
            pady=5,
            command=self._show_help
        )
        help_btn.pack(side=tk.LEFT, padx=5)
        
        # About button
        about_btn = tk.Button(
            buttons_frame,
            text="ℹ️ About",
            font=("Segoe UI", 9),
            foreground="#323130",
            background="#F3F2F1",
            borderwidth=1,
            padx=15,
            pady=5,
            command=self._show_about
        )
        about_btn.pack(side=tk.LEFT, padx=5)
        
        # Exit button
        exit_btn = tk.Button(
            buttons_frame,
            text="🚪 Exit",
            font=("Segoe UI", 9),
            foreground="#FFFFFF",
            background="#D13438",
            borderwidth=0,
            padx=15,
            pady=5,
            command=self._exit_application
        )
        exit_btn.pack(side=tk.LEFT, padx=5)
    
    def _darken_color(self, color):
        """Darken a hex color for hover effects"""
        color_map = {
            "#107C10": "#0E6A0E",
            "#2196F3": "#1976D2",
            "#FF9800": "#F57C00",
            "#0078D4": "#106EBE",
            "#D13438": "#B71C1C"
        }
        return color_map.get(color, color)
    
    def _launch_excel_tool(self):
        """Launch the Excel comparison tool"""
        try:
            # Check if rtest.py exists
            excel_tool_path = Path(__file__).parent / "rtest.py"
            if not excel_tool_path.exists():
                messagebox.showerror(
                    "File Not Found",
                    f"Excel comparison tool not found at:\n{excel_tool_path}\n\n"
                    "Please ensure 'rtest.py' is in the same directory."
                )
                return
            
            # Hide the launcher window
            self.root.withdraw()
            
            # Import and launch the Excel tool
            try:
                import rtest
                excel_root = tk.Toplevel()
                excel_app = rtest.ExcelComparisonApp(excel_root)
                
                # When Excel tool closes, show launcher again
                def on_excel_close():
                    excel_root.destroy()
                    self.root.deiconify()
                    self.root.lift()
                    self.root.focus_force()
                
                excel_root.protocol("WM_DELETE_WINDOW", on_excel_close)
                
            except ImportError as e:
                messagebox.showerror(
                    "Import Error",
                    f"Failed to import Excel comparison tool:\n{str(e)}\n\n"
                    "Please check that all required dependencies are installed."
                )
                self.root.deiconify()
                
        except Exception as e:
            messagebox.showerror("Error", f"Failed to launch Excel tool:\n{str(e)}")
            self.root.deiconify()
    
    def _launch_text_tool(self):
        """Launch the text comparison tool"""
        try:
            # Check if text_compare.py exists
            text_tool_path = Path(__file__).parent / "text_compare.py"
            if not text_tool_path.exists():
                messagebox.showerror(
                    "File Not Found",
                    f"Text comparison tool not found at:\n{text_tool_path}\n\n"
                    "Please ensure 'text_compare.py' is in the same directory."
                )
                return
            
            # Hide the launcher window
            self.root.withdraw()
            
            # Import and launch the text tool
            try:
                import text_compare
                text_root = tk.Toplevel()
                text_app = text_compare.TextComparisonApp(text_root)
                text_app.parent_app = self.root
                
                # When text tool closes, show launcher again
                def on_text_close():
                    text_root.destroy()
                    self.root.deiconify()
                    self.root.lift()
                    self.root.focus_force()
                
                text_root.protocol("WM_DELETE_WINDOW", on_text_close)
                
            except ImportError as e:
                messagebox.showerror(
                    "Import Error",
                    f"Failed to import text comparison tool:\n{str(e)}\n\n"
                    "Please check that all required dependencies are installed."
                )
                self.root.deiconify()
                
        except Exception as e:
            messagebox.showerror("Error", f"Failed to launch text tool:\n{str(e)}")
            self.root.deiconify()
    
    def _launch_image_tool(self):
        """Launch the image comparison tool"""
        try:
            # Check if pc.py exists
            image_tool_path = Path(__file__).parent / "pc.py"
            if not image_tool_path.exists():
                messagebox.showerror(
                    "File Not Found",
                    f"Image comparison tool not found at:\n{image_tool_path}\n\n"
                    "Please ensure 'pc.py' is in the same directory."
                )
                return
            
            # Hide the launcher window
            self.root.withdraw()
            
            # Import and launch the image tool
            try:
                import pc
                image_root = tk.Toplevel()
                image_app = pc.ImageTextComparator(image_root)
                
                # When image tool closes, show launcher again
                def on_image_close():
                    image_root.destroy()
                    self.root.deiconify()
                    self.root.lift()
                    self.root.focus_force()
                
                image_root.protocol("WM_DELETE_WINDOW", on_image_close)
                
            except ImportError as e:
                messagebox.showerror(
                    "Import Error",
                    f"Failed to import image comparison tool:\n{str(e)}\n\n"
                    "Please check that all required dependencies are installed.\n"
                    "Required packages: opencv-python, pytesseract, Pillow, scikit-learn, Levenshtein"
                )
                self.root.deiconify()
                
        except Exception as e:
            messagebox.showerror("Error", f"Failed to launch image tool:\n{str(e)}")
            self.root.deiconify()
    
    def _show_help(self):
        """Show help and documentation"""
        help_text = """
🛠️ Comparison Tools Suite - Help & Documentation

TOOL DESCRIPTIONS:

📊 Excel Compare & Update
- Compare and synchronize Excel files
- Supports both standard and custom matching modes
- Formula-aware processing preserves calculations
- Advanced filtering and bulk operations
- Automatic backup and change highlighting

📄 Text File Comparison  
- Compare any text-based files
- Multiple encoding support
- Side-by-side and unified diff views
- Export to HTML reports
- Advanced search and filtering

🖼️ Image & Picture Comparison
- OCR text extraction (English/Thai)
- Visual difference detection
- Counterfeit document detection
- QR code scanning and validation
- Document authenticity verification

GETTING STARTED:
1. Click 'Launch' on any tool card
2. Follow the tool-specific instructions
3. Use 'Help' buttons within each tool for detailed guidance

SYSTEM REQUIREMENTS:
- Python 3.7 or higher
- Required packages installed (see individual tools)
- Sufficient RAM for large file processing

SUPPORT:
- Use the '❓ Help' buttons in individual tools
- Check file paths and permissions
- Ensure all dependencies are installed
"""
        
        # Create help window
        help_window = tk.Toplevel(self.root)
        help_window.title("Help & Documentation")
        help_window.geometry("700x600")
        help_window.transient(self.root)
        help_window.grab_set()
        
        # Center the help window
        help_window.geometry("+{}+{}".format(
            self.root.winfo_rootx() + 50,
            self.root.winfo_rooty() + 50
        ))
        
        # Create scrollable text area
        text_frame = tk.Frame(help_window)
        text_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=20)
        
        text_widget = tk.Text(
            text_frame,
            wrap=tk.WORD,
            font=("Segoe UI", 10),
            padx=15,
            pady=15
        )
        scrollbar = ttk.Scrollbar(text_frame, orient="vertical", command=text_widget.yview)
        text_widget.configure(yscrollcommand=scrollbar.set)
        
        text_widget.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
        
        # Insert help text
        text_widget.insert("1.0", help_text)
        text_widget.config(state="disabled")
        
        # Close button
        close_btn = tk.Button(
            help_window,
            text="Close",
            command=help_window.destroy,
            padx=20,
            pady=5
        )
        close_btn.pack(pady=10)
    
    def _show_about(self):
        """Show about information"""
        about_text = """
🛠️ Comparison Tools Suite
Version 1.0.0

A comprehensive collection of comparison and analysis tools designed for 
professional data processing and quality assurance workflows.

INCLUDED TOOLS:
• Excel Compare & Update - Advanced spreadsheet comparison
• Text File Comparison - Multi-format text analysis  
• Image & Picture Comparison - OCR and visual analysis

FEATURES:
• Modern, intuitive user interface
• Multi-language support (English, Thai)
• Professional-grade algorithms
• Export and reporting capabilities
• Cross-platform compatibility

DEVELOPED FOR:
Quality Assurance, Data Analysis, Document Verification,
Content Comparison, and File Synchronization tasks.

© 2024 Comparison Tools Suite
All rights reserved.
"""
        
        messagebox.showinfo("About - Comparison Tools Suite", about_text)
    
    def _exit_application(self):
        """Exit the application"""
        if messagebox.askyesno("Exit", "Are you sure you want to exit the Comparison Tools Suite?"):
            self.root.quit()
    
    def _setup_window_state(self):
        """Setup window state management"""
        self.root.protocol("WM_DELETE_WINDOW", self._exit_application)
        
        # Center the window on startup
        self.root.update_idletasks()

def main():
    """Main function to launch the tool launcher"""
    try:
        root = tk.Tk()
        app = ToolLauncher(root)
        root.mainloop()
    except Exception as e:
        print(f"Error starting Tool Launcher: {e}")
        messagebox.showerror("Startup Error", f"Failed to start Tool Launcher:\n{str(e)}")

if __name__ == "__main__":
    main()