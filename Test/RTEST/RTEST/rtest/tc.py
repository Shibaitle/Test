import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import os
import difflib
from pathlib import Path
import threading

class TextComparisonApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Text File Comparison Tool")
        
        # Make the window responsive
        screen_width = root.winfo_screenwidth()
        screen_height = root.winfo_screenheight()
        window_width = int(screen_width * 0.85)
        window_height = int(screen_height * 0.85)
        center_x = int((screen_width - window_width) / 2)
        center_y = int((screen_height - window_height) / 2)
        self.root.geometry(f"{window_width}x{window_height}+{center_x}+{center_y}")
        self.root.minsize(1000, 700)
        
        # Variables to store file paths
        self.old_file_path = tk.StringVar()
        self.new_file_path = tk.StringVar()
        self.file_extension = tk.StringVar(value=".txt")
        
        # Comparison options
        self.ignore_whitespace = tk.BooleanVar(value=False)
        self.ignore_case = tk.BooleanVar(value=False)
        self.show_line_numbers = tk.BooleanVar(value=True)
        self.context_lines = tk.IntVar(value=3)
        
        # Output options
        self.save_diff = tk.BooleanVar(value=False)
        self.diff_format = tk.StringVar(value="unified")  # unified, context, html

        # Encoding options
        self.encoding_mode = tk.StringVar(value="auto")  # auto, same, separate
        self.old_file_encoding = tk.StringVar(value="utf-8")
        self.new_file_encoding = tk.StringVar(value="utf-8")
        self.both_files_encoding = tk.StringVar(value="utf-8")
        
        self._create_ui()
        
        # Initialize encoding widgets
        self._create_encoding_widgets()
        
        # Add proper window close protocol
        self.root.protocol("WM_DELETE_WINDOW", self._on_closing)
    
    def _on_closing(self):
        """Handle window closing event"""
        try:
            if hasattr(self, 'parent_app') and self.parent_app:
                self.parent_app.deiconify()  # Show the parent window
        except:
            pass  # Ignore errors when showing parent
        finally:
            self.root.destroy()

    def _create_ui(self):
        # Update window title to be more descriptive
        self.root.title("📄 Text File Comparison Tool - Excel Compare Tool Extension")
        
        # Add a header with navigation context
        header_frame = ttk.Frame(self.root)
        header_frame.pack(fill=tk.X, padx=10, pady=(10, 5))
        
        # Navigation breadcrumb
        nav_frame = ttk.Frame(header_frame)
        nav_frame.pack(fill=tk.X)
        
        ttk.Label(
            nav_frame,
            text="📊 Excel Compare Tool",
            font=("Arial", 9),
            foreground="#7F8C8D"
        ).pack(side=tk.LEFT)
        
        ttk.Label(
            nav_frame,
            text=" → ",
            font=("Arial", 9),
            foreground="#BDC3C7"
        ).pack(side=tk.LEFT)
        
        ttk.Label(
            nav_frame,
            text="📄 Text File Comparison",
            font=("Arial", 9, "bold"),
            foreground="#2980B9"
        ).pack(side=tk.LEFT)
        
        # Add a subtle separator
        ttk.Separator(header_frame, orient='horizontal').pack(fill=tk.X, pady=(5, 0))
        
        # Create main container with scrollbar
        main_frame = ttk.Frame(self.root)
        main_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # File selection section
        self._create_file_selection(main_frame)
        
        # Options section
        self._create_options_section(main_frame)
        
        # Action buttons
        self._create_action_buttons(main_frame)
        
        # Results section
        self._create_results_section(main_frame)
        
        # Add encoding info display section
        self._create_encoding_info_section(main_frame)
        
        # Status bar
        self._create_status_bar(main_frame)
    
    def _create_file_selection(self, parent):
        file_frame = ttk.LabelFrame(parent, text="File Selection", padding="10")
        file_frame.pack(fill=tk.X, pady=(0, 10))
        
        # File extension selection
        ext_frame = ttk.Frame(file_frame)
        ext_frame.pack(fill=tk.X, pady=(0, 10))
        
        ttk.Label(ext_frame, text="File Extension:").pack(side=tk.LEFT, padx=(0, 5))
        ext_combo = ttk.Combobox(
            ext_frame, 
            textvariable=self.file_extension,
            values=[".txt", ".log", ".csv", ".json", ".xml", ".html", ".py", ".js", ".css", ".md", ".ini", ".cfg", ".*"],
            width=10,
            state="normal"
        )
        ext_combo.pack(side=tk.LEFT, padx=(0, 10))
        
        ttk.Label(ext_frame, text="(Use .* for all files)").pack(side=tk.LEFT)
        
        # Original file selection
        old_frame = ttk.Frame(file_frame)
        old_frame.pack(fill=tk.X, pady=5)
        
        ttk.Label(old_frame, text="Original File:").pack(side=tk.LEFT, padx=(0, 5))
        ttk.Entry(old_frame, textvariable=self.old_file_path, width=60).pack(side=tk.LEFT, padx=(0, 5), fill=tk.X, expand=True)
        ttk.Button(old_frame, text="Browse...", command=self._browse_old_file).pack(side=tk.LEFT)
        
        # New file selection
        new_frame = ttk.Frame(file_frame)
        new_frame.pack(fill=tk.X, pady=5)
        
        ttk.Label(new_frame, text="New File:").pack(side=tk.LEFT, padx=(0, 5))
        ttk.Entry(new_frame, textvariable=self.new_file_path, width=60).pack(side=tk.LEFT, padx=(0, 5), fill=tk.X, expand=True)
        ttk.Button(new_frame, text="Browse...", command=self._browse_new_file).pack(side=tk.LEFT)
    
    def _create_options_section(self, parent):
        options_frame = ttk.LabelFrame(parent, text="Comparison Options", padding="10")
        options_frame.pack(fill=tk.X, pady=(0, 10))
        
        # Comparison options
        comp_frame = ttk.Frame(options_frame)
        comp_frame.pack(fill=tk.X, pady=(0, 5))
        
        ttk.Checkbutton(comp_frame, text="Ignore whitespace", variable=self.ignore_whitespace).pack(side=tk.LEFT, padx=(0, 15))
        ttk.Checkbutton(comp_frame, text="Ignore case", variable=self.ignore_case).pack(side=tk.LEFT, padx=(0, 15))
        ttk.Checkbutton(comp_frame, text="Show line numbers", variable=self.show_line_numbers).pack(side=tk.LEFT, padx=(0, 15))
        
        # Context lines and save options
        context_frame = ttk.Frame(options_frame)
        context_frame.pack(fill=tk.X, pady=5)
        
        ttk.Label(context_frame, text="Context lines:").pack(side=tk.LEFT, padx=(0, 5))
        ttk.Spinbox(context_frame, from_=0, to=10, textvariable=self.context_lines, width=5).pack(side=tk.LEFT, padx=(0, 15))
        
        # Output options
        ttk.Checkbutton(context_frame, text="Save diff to file", variable=self.save_diff).pack(side=tk.LEFT, padx=(0, 15))
        
        ttk.Label(context_frame, text="Diff format:").pack(side=tk.LEFT, padx=(0, 5))
        format_combo = ttk.Combobox(
            context_frame,
            textvariable=self.diff_format,
            values=["unified", "context", "html"],
            width=10,
            state="readonly"
        )
        format_combo.pack(side=tk.LEFT)
        
        # Encoding options section
        encoding_frame = ttk.LabelFrame(options_frame, text="File Encoding Options", padding="5")
        encoding_frame.pack(fill=tk.X, pady=(10, 0))
        
        # Encoding mode selection
        mode_frame = ttk.Frame(encoding_frame)
        mode_frame.pack(fill=tk.X, pady=(0, 5))
        
        ttk.Label(mode_frame, text="Encoding Mode:", font=("Arial", 9, "bold")).pack(side=tk.LEFT, padx=(0, 10))
        
        ttk.Radiobutton(
            mode_frame, 
            text="🔍 Auto-detect", 
            variable=self.encoding_mode, 
            value="auto",
            command=self._on_encoding_mode_change
        ).pack(side=tk.LEFT, padx=(0, 10))
        
        ttk.Radiobutton(
            mode_frame, 
            text="📝 Same for both files", 
            variable=self.encoding_mode, 
            value="same",
            command=self._on_encoding_mode_change
        ).pack(side=tk.LEFT, padx=(0, 10))
        
        ttk.Radiobutton(
            mode_frame, 
            text="⚙️ Separate encodings", 
            variable=self.encoding_mode, 
            value="separate",
            command=self._on_encoding_mode_change
        ).pack(side=tk.LEFT)
        
        # Encoding selection widgets
        self.encoding_widgets_frame = ttk.Frame(encoding_frame)
        self.encoding_widgets_frame.pack(fill=tk.X, pady=(5, 0))
        
        self._create_encoding_widgets()
    
    def _create_encoding_widgets(self):
        """Create encoding selection widgets based on current mode"""
        # Clear existing widgets
        for widget in self.encoding_widgets_frame.winfo_children():
            widget.destroy()
        
        encoding_options = [
            "utf-8", "utf-8-sig", "latin-1", "cp1252", "iso-8859-1", 
            "ascii", "utf-16", "utf-32", "cp437", "cp850", "iso-8859-15"
        ]
        
        mode = self.encoding_mode.get()
        
        if mode == "auto":
            # Show auto-detect info
            info_label = ttk.Label(
                self.encoding_widgets_frame,
                text="ℹ️ Encoding will be automatically detected (utf-8, utf-8-sig, latin-1, cp1252, iso-8859-1)",
                foreground="#666666"
            )
            info_label.pack(anchor=tk.W)
            
        elif mode == "same":
            # Show single encoding selection for both files
            same_frame = ttk.Frame(self.encoding_widgets_frame)
            same_frame.pack(fill=tk.X)
            
            ttk.Label(same_frame, text="📄 Encoding for both files:", font=("Arial", 9, "bold")).pack(side=tk.LEFT, padx=(0, 5))
            
            encoding_combo = ttk.Combobox(
                same_frame,
                textvariable=self.both_files_encoding,
                values=encoding_options,
                width=15,
                state="normal"
            )
            encoding_combo.pack(side=tk.LEFT)
            
        elif mode == "separate":
            # Show separate encoding selection for each file
            old_frame = ttk.Frame(self.encoding_widgets_frame)
            old_frame.pack(fill=tk.X, pady=(0, 3))
            
            ttk.Label(old_frame, text="📄 Original file encoding:", font=("Arial", 9, "bold")).pack(side=tk.LEFT, padx=(0, 5))
            
            old_encoding_combo = ttk.Combobox(
                old_frame,
                textvariable=self.old_file_encoding,
                values=encoding_options,
                width=15,
                state="normal"
            )
            old_encoding_combo.pack(side=tk.LEFT)
            
            new_frame = ttk.Frame(self.encoding_widgets_frame)
            new_frame.pack(fill=tk.X)
            
            ttk.Label(new_frame, text="📄 New file encoding:", font=("Arial", 9, "bold")).pack(side=tk.LEFT, padx=(0, 5))
            
            new_encoding_combo = ttk.Combobox(
                new_frame,
                textvariable=self.new_file_encoding,
                values=encoding_options,
                width=15,
                state="normal"
            )
            new_encoding_combo.pack(side=tk.LEFT)

    def _on_encoding_mode_change(self):
        """Handle encoding mode change"""
        self._create_encoding_widgets()
    
    def _create_action_buttons(self, parent):
        button_frame = ttk.Frame(parent)
        button_frame.pack(fill=tk.X, pady=(0, 10))
        
        # Left side buttons
        left_buttons = ttk.Frame(button_frame)
        left_buttons.pack(side=tk.LEFT)
        
        # Main action buttons with icons
        compare_btn = ttk.Button(
            left_buttons, 
            text="🔍 Compare Files", 
            command=self._start_comparison
        )
        compare_btn.pack(side=tk.LEFT, padx=(0, 10))
        
        clear_btn = ttk.Button(
            left_buttons, 
            text="🗑️ Clear Results", 
            command=self._clear_results
        )
        clear_btn.pack(side=tk.LEFT, padx=(0, 10))
        
        # Right side - Back button (more prominent)
        right_buttons = ttk.Frame(button_frame)
        right_buttons.pack(side=tk.RIGHT)
        
        # More prominent back button
        back_btn = ttk.Button(
            right_buttons,
            text="⬅️ Back to Excel Tool",
            command=self._back_to_excel
        )
        back_btn.pack(side=tk.RIGHT)
        
        # Make the back button more visible
        back_btn.configure(cursor="hand2")
        
        # Add a separator line above buttons for better visual separation
        separator = ttk.Separator(parent, orient='horizontal')
        separator.pack(fill=tk.X, pady=(0, 10))
    
    def _create_results_section(self, parent):
        results_frame = ttk.LabelFrame(parent, text="Comparison Results", padding="10")
        results_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 10))
        
        # Create notebook for different views
        self.notebook = ttk.Notebook(results_frame)
        self.notebook.pack(fill=tk.BOTH, expand=True)
        
        # Summary tab
        summary_frame = ttk.Frame(self.notebook)
        self.notebook.add(summary_frame, text="Summary")
        
        self.summary_text = tk.Text(summary_frame, height=6, wrap=tk.WORD)
        summary_scroll = ttk.Scrollbar(summary_frame, orient=tk.VERTICAL, command=self.summary_text.yview)
        self.summary_text.configure(yscrollcommand=summary_scroll.set)
        
        self.summary_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        summary_scroll.pack(side=tk.RIGHT, fill=tk.Y)
        
        # Unified diff tab
        unified_frame = ttk.Frame(self.notebook)
        self.notebook.add(unified_frame, text="Unified Diff")
        
        self.unified_text = tk.Text(unified_frame, font=("Courier", 10), wrap=tk.NONE)
        unified_scroll_y = ttk.Scrollbar(unified_frame, orient=tk.VERTICAL, command=self.unified_text.yview)
        unified_scroll_x = ttk.Scrollbar(unified_frame, orient=tk.HORIZONTAL, command=self.unified_text.xview)
        self.unified_text.configure(yscrollcommand=unified_scroll_y.set, xscrollcommand=unified_scroll_x.set)
        
        self.unified_text.grid(row=0, column=0, sticky="nsew")
        unified_scroll_y.grid(row=0, column=1, sticky="ns")
        unified_scroll_x.grid(row=1, column=0, sticky="ew")
        
        unified_frame.grid_rowconfigure(0, weight=1)
        unified_frame.grid_columnconfigure(0, weight=1)
        
        # Side-by-side comparison tab
        sidebyside_frame = ttk.Frame(self.notebook)
        self.notebook.add(sidebyside_frame, text="Side by Side")
        
        # Create two text widgets side by side
        left_frame = ttk.LabelFrame(sidebyside_frame, text="Original File")
        left_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(0, 5))
        
        self.left_text = tk.Text(left_frame, font=("Courier", 10), wrap=tk.NONE)
        left_scroll = ttk.Scrollbar(left_frame, orient=tk.VERTICAL, command=self.left_text.yview)
        self.left_text.configure(yscrollcommand=left_scroll.set)
        
        self.left_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        left_scroll.pack(side=tk.RIGHT, fill=tk.Y)
        
        right_frame = ttk.LabelFrame(sidebyside_frame, text="New File")
        right_frame.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True, padx=(5, 0))
        
        self.right_text = tk.Text(right_frame, font=("Courier", 10), wrap=tk.NONE)
        right_scroll = ttk.Scrollbar(right_frame, orient=tk.VERTICAL, command=self.right_text.yview)
        self.right_text.configure(yscrollcommand=right_scroll.set)
        
        self.right_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        right_scroll.pack(side=tk.RIGHT, fill=tk.Y)
        
        # Configure text widget tags for highlighting
        self._configure_text_tags()
    
    def _create_encoding_info_section(self, parent):
        """Create section to display detected/used encodings"""
        self.encoding_info_frame = ttk.LabelFrame(parent, text="File Encoding Information", padding="5")
        self.encoding_info_frame.pack(fill=tk.X, pady=(0, 10))
        
        # Create text widget to display encoding info
        self.encoding_info_text = tk.Text(
            self.encoding_info_frame, 
            height=3, 
            wrap=tk.WORD,
            background="#F8F9FA",
            foreground="#495057",
            font=("Arial", 9),
            state=tk.DISABLED
        )
        self.encoding_info_text.pack(fill=tk.X, pady=5)
        
        # Initially hide the frame
        self.encoding_info_frame.pack_forget()
    
    def _create_status_bar(self, parent):
        self.status_var = tk.StringVar(value="Ready")
        status_frame = ttk.Frame(parent)
        status_frame.pack(fill=tk.X)
        
        ttk.Label(status_frame, text="Status:").pack(side=tk.LEFT)
        ttk.Label(status_frame, textvariable=self.status_var).pack(side=tk.LEFT, padx=(5, 0))
        
        # Add encoding status
        self.encoding_status_var = tk.StringVar(value="")
        self.encoding_status_label = ttk.Label(
            status_frame, 
            textvariable=self.encoding_status_var,
            foreground="#6C757D",
            font=("Arial", 8)
        )
        self.encoding_status_label.pack(side=tk.LEFT, padx=(20, 0))
        
        # Progress bar
        self.progress_var = tk.DoubleVar()
        self.progress_bar = ttk.Progressbar(
            status_frame,
            orient=tk.HORIZONTAL,
            length=200,
            mode='determinate',
            variable=self.progress_var
        )
        self.progress_bar.pack(side=tk.RIGHT, padx=(10, 0))
    
    def _configure_text_tags(self):
        # Configure tags for syntax highlighting
        for text_widget in [self.unified_text, self.left_text, self.right_text]:
            text_widget.tag_configure("added", background="#d4edda", foreground="#155724")
            text_widget.tag_configure("removed", background="#f8d7da", foreground="#721c24")
            text_widget.tag_configure("changed", background="#fff3cd", foreground="#856404")
            text_widget.tag_configure("line_number", foreground="#6c757d", font=("Courier", 8))
    
    def _browse_old_file(self):
        file_types = [("All files", "*.*")]
        ext = self.file_extension.get()
        if ext and ext != ".*":
            file_types.insert(0, (f"{ext.upper()} files", f"*{ext}"))
        
        file_path = filedialog.askopenfilename(
            title="Select Original File",
            filetypes=file_types
        )
        if file_path:
            self.old_file_path.set(file_path)
    
    def _browse_new_file(self):
        file_types = [("All files", "*.*")]
        ext = self.file_extension.get()
        if ext and ext != ".*":
            file_types.insert(0, (f"{ext.upper()} files", f"*{ext}"))
        
        file_path = filedialog.askopenfilename(
            title="Select New File",
            filetypes=file_types
        )
        if file_path:
            self.new_file_path.set(file_path)
    
    def _start_comparison(self):
        if not self._validate_inputs():
            return
        
        # Start comparison in a separate thread to prevent UI freezing
        threading.Thread(target=self._compare_files, daemon=True).start()
    
    def _validate_inputs(self):
        if not self.old_file_path.get():
            messagebox.showerror("Error", "Please select the original file.")
            return False
        
        if not self.new_file_path.get():
            messagebox.showerror("Error", "Please select the new file.")
            return False
        
        if not os.path.exists(self.old_file_path.get()):
            messagebox.showerror("Error", "Original file does not exist.")
            return False
        
        if not os.path.exists(self.new_file_path.get()):
            messagebox.showerror("Error", "New file does not exist.")
            return False
        
        return True
    
    def _compare_files(self):
        try:
            self._update_ui_safely("Reading files...", 10)
            
            old_result = self._read_file(self.old_file_path.get(), "old")
            new_result = self._read_file(self.new_file_path.get(), "new")

            if old_result[0] is None or new_result[0] is None:
                self._update_ui_safely("Error reading files", 0)
                return
            
            old_lines, old_encoding = old_result
            new_lines, new_encoding = new_result
            
            # Store encoding information for display
            self.old_file_encoding_used = old_encoding
            self.new_file_encoding_used = new_encoding
            
            total_lines = len(old_lines) + len(new_lines)
            if total_lines > 50000:  # More than 50k lines total
                result = messagebox.askyesno(
                    "Large File Warning", 
                    f"Files contain {total_lines:,} total lines. "
                    "Comparison may take a while and use significant memory. "
                    "Continue?"
                )
                if not result:
                    self._update_ui_safely("Comparison cancelled", 0)
                    return
            
            self._update_ui_safely("Comparing files...", 30)
            
            # Perform comparison
            differ = difflib.unified_diff(
                old_lines,
                new_lines,
                fromfile=f"Original: {os.path.basename(self.old_file_path.get())}",
                tofile=f"New: {os.path.basename(self.new_file_path.get())}",
                n=self.context_lines.get()
            )
            
            diff_lines = list(differ)
            
            self._update_ui_safely(status="Generating results...", progress=60)
            
            # Generate summary
            self._generate_summary(old_lines, new_lines, diff_lines)
            
            # Display unified diff
            self._display_unified_diff(diff_lines)
            
            # Display side-by-side comparison
            self._display_side_by_side(old_lines, new_lines)
            
            # Save diff if requested
            if self.save_diff.get():
                self._save_diff_file(diff_lines, old_lines, new_lines)
            
            self._update_ui_safely("Comparison complete", 100)
            
        except Exception as e:
            messagebox.showerror("Error", f"An error occurred during comparison: {str(e)}")
            self._update_ui_safely("Error occurred", 0)
    
    def _update_ui_safely(self, status=None, progress=None, encoding_status=None):
        """Safely update UI from background thread"""
        if status:
            self.root.after(0, lambda: self.status_var.set(status))
        if progress is not None:
            self.root.after(0, lambda: self.progress_var.set(progress))
        if encoding_status:
            self.root.after(0, lambda: self.encoding_status_var.set(encoding_status))
    
    def _read_file(self, file_path, file_type="old"):
        """Read file with selected encoding options"""
        try:
            # Check if file exists first
            if not os.path.exists(file_path):
                messagebox.showerror("Error", f"File not found: {file_path}")
                return None, None
            
            mode = self.encoding_mode.get()
            
            if mode == "auto":
                # Auto-detect encoding
                encodings = ['utf-8', 'utf-8-sig', 'latin-1', 'cp1252', 'iso-8859-1']
                
                for encoding in encodings:
                    try:
                        with open(file_path, 'r', encoding=encoding) as f:
                            lines = f.readlines()
                        
                        print(f"Successfully read {file_path} with {encoding} encoding (auto-detected)")
                        return self._process_lines(lines), encoding
                    
                    except UnicodeDecodeError:
                        continue
                    except PermissionError:
                        messagebox.showerror("Error", f"Permission denied accessing file: {file_path}")
                        return None, None
                
                messagebox.showerror("Error", f"Could not read file {file_path}. Unsupported encoding.\nTried: {', '.join(encodings)}")
                return None, None
            
            elif mode == "same":
                # Use same encoding for both files
                encoding = self.both_files_encoding.get()
                return self._read_with_encoding(file_path, encoding)

            elif mode == "separate":
                # Use separate encodings
                if file_type == "old":
                    encoding = self.old_file_encoding.get()
                else:  # new file
                    encoding = self.new_file_encoding.get()
                return self._read_with_encoding(file_path, encoding)
                
        except Exception as e:
            messagebox.showerror("Error", f"Error reading file {file_path}: {str(e)}")
            return None, None

    def _read_with_encoding(self, file_path, encoding):
        """Read file with specific encoding"""
        try:
            with open(file_path, 'r', encoding=encoding) as f:
                lines = f.readlines()
            
            print(f"Successfully read {file_path} with {encoding} encoding")
            return self._process_lines(lines), encoding 
            
        except UnicodeDecodeError as e:
            messagebox.showerror(
                "Encoding Error", 
                f"Could not read {file_path} with {encoding} encoding.\n\n"
                f"Error: {str(e)}\n\n"
                "Try using 'Auto-detect' mode or a different encoding."
            )
            return None, None  
        except PermissionError:
            messagebox.showerror("Error", f"Permission denied accessing file: {file_path}")
            return None, None  
        except Exception as e:
            messagebox.showerror("Error", f"Error reading file {file_path} with {encoding}: {str(e)}")
            return None, None  

    def _process_lines(self, lines):
        """Apply preprocessing options to lines"""
        # Apply preprocessing options
        if self.ignore_whitespace.get():
            lines = [line.strip() + '\n' for line in lines]
        
        if self.ignore_case.get():
            lines = [line.lower() for line in lines]
        
        return lines
    
    def _generate_summary(self, old_lines, new_lines, diff_lines):
        # Calculate statistics
        old_line_count = len(old_lines)
        new_line_count = len(new_lines)
        
        added_lines = sum(1 for line in diff_lines if line.startswith('+') and not line.startswith('+++'))
        removed_lines = sum(1 for line in diff_lines if line.startswith('-') and not line.startswith('---'))
        
        # Get encoding info
        old_encoding = getattr(self, 'old_file_encoding_used', 'Unknown')
        new_encoding = getattr(self, 'new_file_encoding_used', 'Unknown')
        
        # Generate summary text
        summary = f"""File Comparison Summary
    {'=' * 50}

    Original File: {os.path.basename(self.old_file_path.get())}
    - Path: {self.old_file_path.get()}
    - Lines: {old_line_count:,}
    - Size: {self._get_file_size(self.old_file_path.get())}
    - Encoding: {old_encoding}

    New File: {os.path.basename(self.new_file_path.get())}
    - Path: {self.new_file_path.get()}
    - Lines: {new_line_count:,}
    - Size: {self._get_file_size(self.new_file_path.get())}
    - Encoding: {new_encoding}

    Changes:
    - Lines added: {added_lines:,}
    - Lines removed: {removed_lines:,}
    - Net change: {new_line_count - old_line_count:,} lines

    Options Used:
    - Ignore whitespace: {'Yes' if self.ignore_whitespace.get() else 'No'}
    - Ignore case: {'Yes' if self.ignore_case.get() else 'No'}
    - Context lines: {self.context_lines.get()}
    - Encoding mode: {self.encoding_mode.get()}
    """
        
        self.summary_text.config(state=tk.NORMAL)
        self.summary_text.delete(1.0, tk.END)
        self.summary_text.insert(1.0, summary)
        self.summary_text.config(state=tk.DISABLED)
        
        # Update encoding info display
        self._update_encoding_info_display()
    
    def _update_encoding_info_display(self):
        """Update the encoding information display"""
        old_encoding = getattr(self, 'old_file_encoding_used', 'Unknown')
        new_encoding = getattr(self, 'new_file_encoding_used', 'Unknown')
        
        encoding_mode = self.encoding_mode.get()
        mode_text = {
            'auto': 'Auto-detected',
            'same': 'Same for both files',
            'separate': 'Separate encodings'
        }.get(encoding_mode, encoding_mode)
        
        # Create encoding info text
        encoding_info = f"""📄 File Encoding Details:
    Original File: {os.path.basename(self.old_file_path.get())} → {old_encoding}
    New File: {os.path.basename(self.new_file_path.get())} → {new_encoding}
    Mode: {mode_text}"""
        
        # Show the encoding info frame and update text
        self.encoding_info_frame.pack(fill=tk.X, pady=(0, 10))
        
        self.encoding_info_text.config(state=tk.NORMAL)
        self.encoding_info_text.delete(1.0, tk.END)
        self.encoding_info_text.insert(1.0, encoding_info)
        
        # Color code based on encoding match
        if old_encoding == new_encoding:
            # Same encoding - green background
            self.encoding_info_text.config(background="#D4EDDA", foreground="#155724")
        else:
            # Different encodings - yellow background
            self.encoding_info_text.config(background="#FFF3CD", foreground="#856404")
        
        self.encoding_info_text.config(state=tk.DISABLED)
    
    def _get_file_size(self, file_path):
        try:
            if not os.path.exists(file_path):
                return "File not found"
            size = os.path.getsize(file_path)
            if size < 1024:
                return f"{size} bytes"
            elif size < 1024 * 1024:
                return f"{size / 1024:.1f} KB"
            else:
                return f"{size / (1024 * 1024):.1f} MB"
        except (OSError, PermissionError) as e:
            return f"Error: {str(e)}"
        except Exception:
            return "Unknown"
    
    def _display_unified_diff(self, diff_lines):
        self.unified_text.config(state=tk.NORMAL)
        self.unified_text.delete(1.0, tk.END)
        
        if not diff_lines:
            self.unified_text.insert(tk.END, "No differences found between the files.")
        else:
            for line in diff_lines:
                start_pos = self.unified_text.index(tk.INSERT)
                self.unified_text.insert(tk.END, line)
                end_pos = self.unified_text.index(tk.INSERT)
                
                # Apply syntax highlighting
                if line.startswith('+') and not line.startswith('+++'):
                    self.unified_text.tag_add("added", start_pos, end_pos)
                elif line.startswith('-') and not line.startswith('---'):
                    self.unified_text.tag_add("removed", start_pos, end_pos)
                elif line.startswith('@@'):
                    self.unified_text.tag_add("changed", start_pos, end_pos)
        
        self.unified_text.config(state=tk.DISABLED)
    
    def _display_side_by_side(self, old_lines, new_lines):
        # Clear both text widgets
        self.left_text.config(state=tk.NORMAL)
        self.right_text.config(state=tk.NORMAL)
        self.left_text.delete(1.0, tk.END)
        self.right_text.delete(1.0, tk.END)
        
        # Create sequence matcher for line-by-line comparison
        matcher = difflib.SequenceMatcher(None, old_lines, new_lines)
        
        old_line_num = 1
        new_line_num = 1
        
        for tag, i1, i2, j1, j2 in matcher.get_opcodes():
            if tag == 'equal':
                # Lines are the same
                for i in range(i1, i2):
                    line_content = old_lines[i] if i < len(old_lines) else ""
                    if self.show_line_numbers.get():
                        self.left_text.insert(tk.END, f"{old_line_num:4d}: ")
                        self.right_text.insert(tk.END, f"{new_line_num:4d}: ")
                        old_line_num += 1
                        new_line_num += 1
                    self.left_text.insert(tk.END, line_content)
                    self.right_text.insert(tk.END, line_content)
            
            elif tag == 'delete':
                # Lines removed from old file
                for i in range(i1, i2):
                    line_content = old_lines[i] if i < len(old_lines) else ""
                    start_pos_left = self.left_text.index(tk.INSERT)
                    if self.show_line_numbers.get():
                        self.left_text.insert(tk.END, f"{old_line_num:4d}: ")
                        self.right_text.insert(tk.END, "     ")
                        old_line_num += 1
                    self.left_text.insert(tk.END, line_content)
                    self.right_text.insert(tk.END, "\n")
                    end_pos_left = self.left_text.index(tk.INSERT)
                    self.left_text.tag_add("removed", start_pos_left, end_pos_left)
            
            elif tag == 'insert':
                # Lines added to new file
                for j in range(j1, j2):
                    line_content = new_lines[j] if j < len(new_lines) else ""
                    start_pos_right = self.right_text.index(tk.INSERT)
                    if self.show_line_numbers.get():
                        self.left_text.insert(tk.END, "     ")
                        self.right_text.insert(tk.END, f"{new_line_num:4d}: ")
                        new_line_num += 1
                    self.left_text.insert(tk.END, "\n")
                    self.right_text.insert(tk.END, line_content)
                    end_pos_right = self.right_text.index(tk.INSERT)
                    self.right_text.tag_add("added", start_pos_right, end_pos_right)
            
            elif tag == 'replace':
                # Lines changed
                max_lines = max(i2 - i1, j2 - j1)
                for k in range(max_lines):
                    old_line = old_lines[i1 + k] if (i1 + k) < i2 and (i1 + k) < len(old_lines) else "\n"
                    new_line = new_lines[j1 + k] if (j1 + k) < j2 and (j1 + k) < len(new_lines) else "\n"
                    
                    # Left side (old line)
                    start_pos_left = self.left_text.index(tk.INSERT)
                    if self.show_line_numbers.get() and (i1 + k) < i2:
                        self.left_text.insert(tk.END, f"{old_line_num:4d}: ")
                        old_line_num += 1
                    elif self.show_line_numbers.get():
                        self.left_text.insert(tk.END, "     ")
                    
                    if (i1 + k) < i2:
                        self.left_text.insert(tk.END, old_line)
                    else:
                        self.left_text.insert(tk.END, "\n")
                    end_pos_left = self.left_text.index(tk.INSERT)
                    
                    # Right side (new line)
                    start_pos_right = self.right_text.index(tk.INSERT)
                    if self.show_line_numbers.get() and (j1 + k) < j2:
                        self.right_text.insert(tk.END, f"{new_line_num:4d}: ")
                        new_line_num += 1
                    elif self.show_line_numbers.get():
                        self.right_text.insert(tk.END, "     ")
                    
                    if (j1 + k) < j2:
                        self.right_text.insert(tk.END, new_line)
                    else:
                        self.right_text.insert(tk.END, "\n")
                    end_pos_right = self.right_text.index(tk.INSERT)
                    
                    # Apply highlighting
                    if (i1 + k) < i2:
                        self.left_text.tag_add("removed", start_pos_left, end_pos_left)
                    if (j1 + k) < j2:
                        self.right_text.tag_add("added", start_pos_right, end_pos_right)
        
        self.left_text.config(state=tk.DISABLED)
        self.right_text.config(state=tk.DISABLED)
    
    def _save_diff_file(self, diff_lines, old_lines, new_lines):
        try:
            # Generate filename based on original files
            old_name = Path(self.old_file_path.get()).stem
            new_name = Path(self.new_file_path.get()).stem
            
            if self.diff_format.get() == "html":
                output_file = filedialog.asksaveasfilename(
                    title="Save Diff Report",
                    defaultextension=".html",
                    filetypes=[("HTML files", "*.html"), ("All files", "*.*")],
                    initialfile=f"diff_{old_name}_vs_{new_name}.html"
                )
                if output_file:
                    self._save_html_diff(output_file, old_lines, new_lines)
            else:
                output_file = filedialog.asksaveasfilename(
                    title="Save Diff Report",
                    defaultextension=".txt",
                    filetypes=[("Text files", "*.txt"), ("All files", "*.*")],
                    initialfile=f"diff_{old_name}_vs_{new_name}.txt"
                )
                if output_file:
                    self._save_text_diff(output_file, diff_lines)
            
            if output_file:
                messagebox.showinfo("Saved", f"Diff report saved to:\n{output_file}")
                
        except Exception as e:
            messagebox.showerror("Error", f"Failed to save diff file: {str(e)}")
    
    def _save_text_diff(self, output_file, diff_lines):
        with open(output_file, 'w', encoding='utf-8') as f:
            for line in diff_lines:
                f.write(line)
    
    def _save_html_diff(self, output_file, old_lines, new_lines):
        # Create HTML diff using difflib
        differ = difflib.HtmlDiff()
        html_diff = differ.make_file(
            old_lines,
            new_lines,
            fromdesc=f"Original: {os.path.basename(self.old_file_path.get())}",
            todesc=f"New: {os.path.basename(self.new_file_path.get())}",
            context=self.context_lines.get() > 0,
            numlines=self.context_lines.get()
        )
        
        with open(output_file, 'w', encoding='utf-8') as f:
            f.write(html_diff)
    
    def _clear_results(self):
        self.summary_text.config(state=tk.NORMAL)
        self.summary_text.delete(1.0, tk.END)
        self.summary_text.config(state=tk.DISABLED)
        
        self.unified_text.config(state=tk.NORMAL)
        self.unified_text.delete(1.0, tk.END)
        self.unified_text.config(state=tk.DISABLED)
        
        self.left_text.config(state=tk.NORMAL)
        self.right_text.config(state=tk.NORMAL)
        self.left_text.delete(1.0, tk.END)
        self.right_text.delete(1.0, tk.END)
        self.left_text.config(state=tk.DISABLED)
        self.right_text.config(state=tk.DISABLED)
        
        # Hide encoding info
        self.encoding_info_frame.pack_forget()
        
        # Clear stored encoding info
        if hasattr(self, 'old_file_encoding_used'):
            delattr(self, 'old_file_encoding_used')
        if hasattr(self, 'new_file_encoding_used'):
            delattr(self, 'new_file_encoding_used')
        
        self.status_var.set("Ready")
        self.progress_var.set(0)
    
    def _back_to_excel(self):
        """Return to the main Excel comparison tool"""
        try:
            if hasattr(self, 'parent_app') and self.parent_app:
                self.parent_app.deiconify()  # Show the parent window
            else:
                # If no parent app, just close this window
                print("No parent app found, closing text comparison window")
            self.root.destroy()
        except Exception as e:
            print(f"Error returning to Excel tool: {e}")
            self.root.destroy()  # Still close the window

    def set_parent_launcher(self, parent_launcher):
        """Set reference to parent launcher for proper window management"""
        self.parent_launcher = parent_launcher

def main():
    root = tk.Tk()
    app = TextComparisonApp(root)
    root.mainloop()

if __name__ == "__main__":
    main()