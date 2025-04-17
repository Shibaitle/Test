import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import cv2
import pytesseract
import Levenshtein
import os
import numpy as np
from sklearn.cluster import DBSCAN

import re
from datetime import datetime
import cv2.aruco as aruco

# Handle Pillow deprecation warning for LANCZOS
from PIL import Image, ImageTk
try:
    # For newer Pillow versions (9.0.0 and above)
    LANCZOS = Image.Resampling.LANCZOS
except AttributeError:
    # For older Pillow versions
    LANCZOS = Image.LANCZOS

class ImageTextComparator:
    def __init__(self, root):
        self.root = root
        self.root.title("Image Text Comparison Tool")
        self.root.geometry("950x700")  # Increased size for better viewing
        self.root.configure(bg="#f0f0f0")
        
        # Initialize variables
        self.image1_path = None
        self.image2_path = None
        self.image1_text = ""
        self.image2_text = ""
        self.image1_cv = None  # Original CV2 image
        self.image2_cv = None  # Original CV2 image
        
        # Language selection variables
        self.lang1_var = tk.StringVar(value="eng")  # Default to English for image 1
        self.lang2_var = tk.StringVar(value="eng")  # Default to English for image 2
        
        # Create notebook for tabs
        self.notebook = ttk.Notebook(root)
        self.notebook.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # Tab 1: Text comparison
        self.text_tab = tk.Frame(self.notebook, bg="#f0f0f0")
        self.notebook.add(self.text_tab, text="Text Comparison")
        
        # Tab 2: Visual comparison
        self.visual_tab = tk.Frame(self.notebook, bg="#f0f0f0")
        self.notebook.add(self.visual_tab, text="Visual Comparison")
        
        # Tab 3: Counterfeit Detection (new)
        self.detection_tab = tk.Frame(self.notebook, bg="#f0f0f0")
        self.notebook.add(self.detection_tab, text="Advanced Detection")
        
        # Tab 4: Bill/Slip Verification
        self.bill_tab = tk.Frame(self.notebook, bg="#f0f0f0")
        self.notebook.add(self.bill_tab, text="Bill Verification")
        
        # Setup text comparison tab
        self.setup_text_tab()
        
        # Setup visual comparison tab
        self.setup_visual_tab()
        
        # Setup counterfeit detection tab (new)
        self.setup_detection_tab()
        
        # Setup bill verification tab
        self.setup_bill_verification_tab()
        
    def setup_text_tab(self):
        """Setup the text comparison tab"""
        # Create main frame
        self.main_frame = tk.Frame(self.text_tab, bg="#f0f0f0")
        self.main_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=20)
        
        # Create frames for images
        self.image_frames_container = tk.Frame(self.main_frame, bg="#f0f0f0")
        self.image_frames_container.pack(fill=tk.BOTH, expand=True)
        
        # Create left frame (Image 1)
        self.left_frame = tk.LabelFrame(self.image_frames_container, text="Image 1", bg="#f0f0f0", padx=10, pady=10)
        self.left_frame.grid(row=0, column=0, padx=10, pady=10, sticky="nsew")
        
        # Create right frame (Image 2)
        self.right_frame = tk.LabelFrame(self.image_frames_container, text="Image 2", bg="#f0f0f0", padx=10, pady=10)
        self.right_frame.grid(row=0, column=1, padx=10, pady=10, sticky="nsew")
        
        # Configure grid weights
        self.image_frames_container.grid_columnconfigure(0, weight=1)
        self.image_frames_container.grid_columnconfigure(1, weight=1)
        self.image_frames_container.grid_rowconfigure(0, weight=1)
        
        # Create image labels
        self.image1_label = tk.Label(self.left_frame, text="No image selected", bg="#e0e0e0", width=40, height=15)
        self.image1_label.pack(fill=tk.BOTH, expand=True)
        
        self.image2_label = tk.Label(self.right_frame, text="No image selected", bg="#e0e0e0", width=40, height=15)
        self.image2_label.pack(fill=tk.BOTH, expand=True)
        
        # Control panel for image 1
        self.control_frame1 = tk.Frame(self.left_frame, bg="#f0f0f0")
        self.control_frame1.pack(fill=tk.X, pady=5)
        
        # Browse button for image 1
        self.browse_btn1 = tk.Button(self.control_frame1, text="Browse", command=lambda: self.browse_image(1))
        self.browse_btn1.pack(side=tk.LEFT, padx=5)
        
        # Language selection for image 1 (centered)
        self.lang_frame1 = tk.Frame(self.control_frame1, bg="#f0f0f0")
        self.lang_frame1.pack(side=tk.LEFT, fill=tk.X, expand=True)
        
        self.lang_label1 = tk.Label(self.lang_frame1, text="Language:", bg="#f0f0f0")
        self.lang_label1.pack(side=tk.LEFT)
        
        # Language dropdown for image 1
        self.lang1_var = tk.StringVar(value="eng")
        self.lang1_dropdown = ttk.Combobox(self.lang_frame1, textvariable=self.lang1_var, 
                                           values=["eng", "tha", "tha+eng"], width=8, state="readonly")
        self.lang1_dropdown.pack(side=tk.LEFT, padx=5)
        
        # Extract text button for image 1
        self.extract_btn1 = tk.Button(self.control_frame1, text="Extract Text", 
                                     command=lambda: self.extract_and_display(1))
        self.extract_btn1.pack(side=tk.RIGHT, padx=5)
        
        # Control panel for image 2
        self.control_frame2 = tk.Frame(self.right_frame, bg="#f0f0f0")
        self.control_frame2.pack(fill=tk.X, pady=5)
        
        # Browse button for image 2
        self.browse_btn2 = tk.Button(self.control_frame2, text="Browse", command=lambda: self.browse_image(2))
        self.browse_btn2.pack(side=tk.LEFT, padx=5)
        
        # Language selection for image 2 (centered)
        self.lang_frame2 = tk.Frame(self.control_frame2, bg="#f0f0f0")
        self.lang_frame2.pack(side=tk.LEFT, fill=tk.X, expand=True)
        
        self.lang_label2 = tk.Label(self.lang_frame2, text="Language:", bg="#f0f0f0")
        self.lang_label2.pack(side=tk.LEFT)
        
        # Language dropdown for image 2
        self.lang2_var = tk.StringVar(value="eng")
        self.lang2_dropdown = ttk.Combobox(self.lang_frame2, textvariable=self.lang2_var, 
                                          values=["eng", "tha", "tha+eng"], width=8, state="readonly")
        self.lang2_dropdown.pack(side=tk.LEFT, padx=5)
        
        # Extract text button for image 2
        self.extract_btn2 = tk.Button(self.control_frame2, text="Extract Text", 
                                     command=lambda: self.extract_and_display(2))
        self.extract_btn2.pack(side=tk.RIGHT, padx=5)
        
        # Text display areas
        self.text_frame = tk.Frame(self.main_frame, bg="#f0f0f0")
        self.text_frame.pack(fill=tk.BOTH, expand=True, pady=5)
        
        # Text area for Image 1 with scrollbar
        self.text_frame1 = tk.LabelFrame(self.text_frame, text="Text from Image 1", bg="#f0f0f0")
        self.text_frame1.grid(row=0, column=0, padx=5, pady=5, sticky="nsew")
        
        # Create a frame to hold the text widget and scrollbar
        self.text1_container = tk.Frame(self.text_frame1)
        self.text1_container.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        # Add the text widget and scrollbars
        self.text1 = tk.Text(self.text1_container, height=5, width=40)
        self.text1_vscroll = ttk.Scrollbar(self.text1_container, orient=tk.VERTICAL, command=self.text1.yview)
        self.text1.configure(yscrollcommand=self.text1_vscroll.set)
        
        # Pack the text widget and scrollbar
        self.text1.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        self.text1_vscroll.pack(side=tk.RIGHT, fill=tk.Y)
        
        # Same for image 2
        self.text_frame2 = tk.LabelFrame(self.text_frame, text="Text from Image 2", bg="#f0f0f0")
        self.text_frame2.grid(row=0, column=1, padx=5, pady=5, sticky="nsew")
        
        self.text2_container = tk.Frame(self.text_frame2)
        self.text2_container.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        self.text2 = tk.Text(self.text2_container, height=5, width=40)
        self.text2_vscroll = ttk.Scrollbar(self.text2_container, orient=tk.VERTICAL, command=self.text2.yview)
        self.text2.configure(yscrollcommand=self.text2_vscroll.set)
        
        self.text2.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        self.text2_vscroll.pack(side=tk.RIGHT, fill=tk.Y)
        
        # Configure grid weights for text frames
        self.text_frame.grid_columnconfigure(0, weight=1)
        self.text_frame.grid_columnconfigure(1, weight=1)
        self.text_frame.grid_rowconfigure(0, weight=1)
        
        # Action buttons
        self.action_frame = tk.Frame(self.main_frame, bg="#f0f0f0")
        self.action_frame.pack(fill=tk.X, pady=10)
        
        # Compare button
        self.compare_btn = tk.Button(self.action_frame, text="Compare Text", 
                                   command=self.compare_text, bg="#4CAF50", fg="white", padx=10, pady=5)
        self.compare_btn.pack(side=tk.LEFT, padx=5)
        
        # Visual diff button
        self.visual_diff_btn = tk.Button(self.action_frame, text="Show Visual Diff", 
                                       command=self.show_visual_diff, bg="#2196F3", fg="white", padx=10, pady=5)
        self.visual_diff_btn.pack(side=tk.LEFT, padx=5)
        
        # Reset button
        self.reset_btn = tk.Button(self.action_frame, text="Reset", 
                                 command=self.reset, bg="#f44336", fg="white", padx=10, pady=5)
        self.reset_btn.pack(side=tk.LEFT, padx=5)
        
        # Result frame
        self.result_frame = tk.LabelFrame(self.main_frame, text="Comparison Result", bg="#f0f0f0", padx=10, pady=10)
        self.result_frame.pack(fill=tk.X, pady=10)
        
        self.result_label = tk.Label(self.result_frame, text="No comparison performed yet", bg="#f0f0f0", font=("Arial", 12))
        self.result_label.pack(pady=10)
        
    def setup_visual_tab(self):
        """Setup the visual comparison tab"""
        # Create main frame for visual comparison
        self.visual_main_frame = tk.Frame(self.visual_tab, bg="#f0f0f0")
        self.visual_main_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=20)
        
        # Create frame for visual difference
        self.visual_diff_frame = tk.LabelFrame(self.visual_main_frame, text="Visual Difference (Zoomable)", bg="#f0f0f0", padx=10, pady=10)
        self.visual_diff_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # Create zoomable canvas for the comparison image
        self.visual_canvas = ZoomableCanvas(self.visual_diff_frame, bg="#000000")
        self.visual_canvas.pack(fill=tk.BOTH, expand=True)
        
        # Set default message
        self.visual_canvas.set_text("No visual comparison available\nClick 'Show Visual Diff' button to generate")
        
        # Instructions
        self.visual_instructions = tk.Label(self.visual_main_frame, 
                                         text="Red areas indicate differences between the images.\n" +
                                         "Use mouse wheel or +/- buttons to zoom. Drag to pan.\n" +
                                         "For best results, use similar-sized images with similar content.",
                                         bg="#f0f0f0", font=("Arial", 10))
        self.visual_instructions.pack(pady=10)

    def setup_detection_tab(self):
        """Setup the counterfeit detection tab with bounding boxes to highlight differences"""
        # Create main frame for detection
        self.detection_main_frame = tk.Frame(self.detection_tab, bg="#f0f0f0")
        self.detection_main_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=20)
        
        # Top control panel for detection settings
        self.detection_control_frame = tk.LabelFrame(self.detection_main_frame, text="Detection Controls", bg="#f0f0f0", padx=10, pady=10)
        self.detection_control_frame.pack(fill=tk.X, pady=10)
        
        # Sensitivity slider for detection
        self.sensitivity_frame = tk.Frame(self.detection_control_frame, bg="#f0f0f0")
        self.sensitivity_frame.pack(side=tk.LEFT, padx=10, pady=5)
        self.sensitivity_label = tk.Label(self.sensitivity_frame, text="Detection Sensitivity:", bg="#f0f0f0")
        self.sensitivity_label.pack(side=tk.LEFT)
        self.sensitivity_var = tk.DoubleVar(value=50.0)  # Default to 50%
        self.sensitivity_scale = ttk.Scale(self.sensitivity_frame, from_=1, to=100, variable=self.sensitivity_var, 
                                          orient=tk.HORIZONTAL, length=200)
        self.sensitivity_scale.pack(side=tk.LEFT, padx=5)
        self.sensitivity_value = tk.Label(self.sensitivity_frame, text="50%", width=5, bg="#f0f0f0")
        self.sensitivity_value.pack(side=tk.LEFT)
        
        # Detection method selection
        self.method_frame = tk.Frame(self.detection_control_frame, bg="#f0f0f0")
        self.method_frame.pack(side=tk.LEFT, padx=20, pady=5)
        self.method_label = tk.Label(self.method_frame, text="Detection Method:", bg="#f0f0f0")
        self.method_label.pack(side=tk.LEFT)
        self.method_var = tk.StringVar(value="contour")  # Default to contour detection
        methods = [("Contour", "contour"), ("Feature Match", "feature")]  # Removed color analysis
        for text, value in methods:
            tk.Radiobutton(self.method_frame, text=text, variable=self.method_var, value=value, bg="#f0f0f0").pack(side=tk.LEFT)
        
        # Detect button
        self.detect_btn = tk.Button(self.detection_control_frame, text="Detect Differences", 
                                  command=self.detect_counterfeit, bg="#FF9800", fg="white", padx=10, pady=5)
        self.detect_btn.pack(side=tk.RIGHT, padx=10)
        
        # Create a frame for the detection result view
        self.detection_view_frame = tk.LabelFrame(self.detection_main_frame, text="Detection Results", 
                                               bg="#f0f0f0", padx=10, pady=10)
        self.detection_view_frame.pack(fill=tk.BOTH, expand=True, pady=10)
        
        # Create zoomable canvas for the detection visualization
        self.detection_canvas = ZoomableCanvas(self.detection_view_frame, bg="#000000")
        self.detection_canvas.pack(fill=tk.BOTH, expand=True)
        
        # Set default message
        self.detection_canvas.set_text("No detection analysis performed\nClick 'Detect Differences' to analyze images")
        
        # Create frame for detailed results
        self.detection_details_frame = tk.LabelFrame(self.detection_main_frame, text="Detection Details", 
                                                  bg="#f0f0f0", padx=10, pady=10)
        self.detection_details_frame.pack(fill=tk.X, pady=10)
        
        # Text area for detection details with scrollbar
        self.details_container = tk.Frame(self.detection_details_frame)
        self.details_container.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        self.detection_details = tk.Text(self.details_container, height=5, width=40)
        self.details_vscroll = ttk.Scrollbar(self.details_container, orient=tk.VERTICAL, command=self.detection_details.yview)
        self.detection_details.configure(yscrollcommand=self.details_vscroll.set)
        
        self.detection_details.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        self.details_vscroll.pack(side=tk.RIGHT, fill=tk.Y)
        
        # Update sensitivity value display when slider changes
        self.sensitivity_scale.config(command=self.update_sensitivity_display)
        
        # Detection instructions
        self.detection_instructions = tk.Label(self.detection_main_frame, 
                                            text="This tool highlights suspicious areas in the images.\n"
                                            "Red boxes indicate potential counterfeits or manipulations.\n"
                                            "Adjust sensitivity to fine-tune detection.",
                                            bg="#f0f0f0", font=("Arial", 10))
        self.detection_instructions.pack(pady=5)

    def setup_bill_verification_tab(self):
        """Setup the bill/slip verification tab"""
        # Create main frame for bill verification
        self.bill_main_frame = tk.Frame(self.bill_tab, bg="#f0f0f0")
        self.bill_main_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=20)
        
        # Top section with type selection and upload
        self.bill_type_frame = tk.LabelFrame(self.bill_main_frame, text="Document Type", bg="#f0f0f0", padx=10, pady=10)
        self.bill_type_frame.pack(fill=tk.X, pady=10)
        
        # Document type selection
        self.bill_type_var = tk.StringVar(value="banking")
        tk.Radiobutton(self.bill_type_frame, text="Banking Slip", variable=self.bill_type_var, 
                       value="banking", bg="#f0f0f0").pack(side=tk.LEFT, padx=5)
        tk.Radiobutton(self.bill_type_frame, text="Payment Receipt", variable=self.bill_type_var, 
                       value="receipt", bg="#f0f0f0").pack(side=tk.LEFT, padx=5)
        tk.Radiobutton(self.bill_type_frame, text="QR Payment", variable=self.bill_type_var, 
                       value="qr", bg="#f0f0f0").pack(side=tk.LEFT, padx=5)
        
        # Browse button
        self.bill_browse_btn = tk.Button(self.bill_type_frame, text="Upload Document", 
                                       command=self.browse_bill, bg="#4CAF50", fg="white", padx=10)
        self.bill_browse_btn.pack(side=tk.RIGHT, padx=10)
        
        # Middle section with key verification fields
        self.bill_verify_frame = tk.LabelFrame(self.bill_main_frame, text="Verification Parameters", bg="#f0f0f0", padx=10, pady=10)
        self.bill_verify_frame.pack(fill=tk.X, pady=10)
        
        # Transaction ID verification
        self.tx_id_frame = tk.Frame(self.bill_verify_frame, bg="#f0f0f0")
        self.tx_id_frame.pack(fill=tk.X, pady=5)
        tk.Label(self.tx_id_frame, text="Transaction ID:", bg="#f0f0f0", width=15, anchor="w").pack(side=tk.LEFT)
        self.tx_id_var = tk.StringVar()
        self.tx_id_entry = tk.Entry(self.tx_id_frame, textvariable=self.tx_id_var, width=40)
        self.tx_id_entry.pack(side=tk.LEFT, padx=5, fill=tk.X, expand=True)
        self.tx_id_auto_btn = tk.Button(self.tx_id_frame, text="Auto Extract", command=lambda: self.auto_extract("tx_id"))
        self.tx_id_auto_btn.pack(side=tk.RIGHT, padx=5)
        
        # Amount verification
        self.amount_frame = tk.Frame(self.bill_verify_frame, bg="#f0f0f0")
        self.amount_frame.pack(fill=tk.X, pady=5)
        tk.Label(self.amount_frame, text="Amount:", bg="#f0f0f0", width=15, anchor="w").pack(side=tk.LEFT)
        self.amount_var = tk.StringVar()
        self.amount_entry = tk.Entry(self.amount_frame, textvariable=self.amount_var, width=40)
        self.amount_entry.pack(side=tk.LEFT, padx=5, fill=tk.X, expand=True)
        self.amount_auto_btn = tk.Button(self.amount_frame, text="Auto Extract", command=lambda: self.auto_extract("amount"))
        self.amount_auto_btn.pack(side=tk.RIGHT, padx=5)
        
        # Date/Time verification
        self.datetime_frame = tk.Frame(self.bill_verify_frame, bg="#f0f0f0")
        self.datetime_frame.pack(fill=tk.X, pady=5)
        tk.Label(self.datetime_frame, text="Date/Time:", bg="#f0f0f0", width=15, anchor="w").pack(side=tk.LEFT)
        self.datetime_var = tk.StringVar()
        self.datetime_entry = tk.Entry(self.datetime_frame, textvariable=self.datetime_var, width=40)
        self.datetime_entry.pack(side=tk.LEFT, padx=5, fill=tk.X, expand=True)
        self.datetime_auto_btn = tk.Button(self.datetime_frame, text="Auto Extract", command=lambda: self.auto_extract("datetime"))
        self.datetime_auto_btn.pack(side=tk.RIGHT, padx=5)
        
        # Action buttons
        self.bill_action_frame = tk.Frame(self.bill_main_frame, bg="#f0f0f0")
        self.bill_action_frame.pack(fill=tk.X, pady=10)
        
        self.verify_btn = tk.Button(self.bill_action_frame, text="Verify Document", 
                                   command=self.verify_bill, bg="#2196F3", fg="white", padx=10, pady=5)
        self.verify_btn.pack(side=tk.LEFT, padx=5)
        
        self.scan_qr_btn = tk.Button(self.bill_action_frame, text="Scan QR Code", 
                                    command=self.scan_qr_code, bg="#9C27B0", fg="white", padx=10, pady=5)
        self.scan_qr_btn.pack(side=tk.LEFT, padx=5)
        
        self.bill_reset_btn = tk.Button(self.bill_action_frame, text="Reset", 
                                       command=self.reset_bill_verification, bg="#f44336", fg="white", padx=10, pady=5)
        self.bill_reset_btn.pack(side=tk.LEFT, padx=5)
        
        # Display area for the document
        self.bill_display_frame = tk.LabelFrame(self.bill_main_frame, text="Document", bg="#f0f0f0", padx=10, pady=10)
        self.bill_display_frame.pack(fill=tk.BOTH, expand=True, pady=10)
        
        self.bill_canvas = ZoomableCanvas(self.bill_display_frame, bg="#000000")
        self.bill_canvas.pack(fill=tk.BOTH, expand=True)
        self.bill_canvas.set_text("No document uploaded\nUse 'Upload Document' to begin verification")
        
        # Results section
        self.bill_results_frame = tk.LabelFrame(self.bill_main_frame, text="Verification Results", bg="#f0f0f0", padx=10, pady=10)
        self.bill_results_frame.pack(fill=tk.X, pady=10)
        
        self.results_text = tk.Text(self.bill_results_frame, height=8, width=80, wrap=tk.WORD)
        self.results_text.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        self.results_text.config(state=tk.DISABLED)
        
        # Initialize bill image variable
        self.bill_path = None
        self.bill_cv = None

    def update_sensitivity_display(self, value=None):
        """Update the sensitivity value display"""
        self.sensitivity_value.config(text=f"{int(self.sensitivity_var.get())}%")

    def browse_image(self, image_num):
        """Open a file dialog to select an image"""
        file_path = filedialog.askopenfilename(
            title=f"Select Image {image_num}",
            filetypes=(("Image files", "*.jpg;*.jpeg;*.png;*.bmp"), ("All files", "*.*"))
        )
        
        if file_path:
            try:
                # Store the original image for processing
                cv_img = cv2.imdecode(np.fromfile(file_path, dtype=np.uint8), cv2.IMREAD_UNCHANGED)
                
                # Load and resize image for display
                img = Image.open(file_path)
                img_aspect = img.width / img.height
                
                # Calculate new dimensions to fit in label while maintaining aspect ratio
                max_width = 350
                max_height = 250
                
                if img_aspect > 1:  # Wider than tall
                    new_width = min(img.width, max_width)
                    new_height = int(new_width / img_aspect)
                else:  # Taller than wide
                    new_height = min(img.height, max_height)
                    new_width = int(new_height * img_aspect)
                
                # Use the version-safe LANCZOS constant
                img = img.resize((new_width, new_height), LANCZOS)
                photo = ImageTk.PhotoImage(img)
                
                if image_num == 1:
                    self.image1_path = file_path
                    self.image1_cv = cv_img  # Store original CV2 image
                    self.image1_label.config(image=photo)
                    self.image1_label.image = photo  # Keep a reference
                else:
                    self.image2_path = file_path
                    self.image2_cv = cv_img  # Store original CV2 image
                    self.image2_label.config(image=photo)
                    self.image2_label.image = photo  # Keep a reference
                
            except Exception as e:
                messagebox.showerror("Error", f"Failed to open image: {e}")
    
    def extract_text(self, image_path, lang="eng"):
        """Extract text from an image using pytesseract OCR with enhanced preprocessing"""
        if not image_path:
            return ""
        
        try:
            # Read image with OpenCV - handle non-ASCII filenames
            img = cv2.imdecode(np.fromfile(image_path, dtype=np.uint8), cv2.IMREAD_UNCHANGED)
            
            # Convert to grayscale if needed
            if len(img.shape) == 3:
                gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
            else:
                gray = img.copy()
            
            # Resize for better OCR if image is small
            height, width = gray.shape
            if width < 1000 or height < 1000:
                scale_factor = 2.0
                gray = cv2.resize(gray, None, fx=scale_factor, fy=scale_factor, interpolation=cv2.INTER_CUBIC)
            
            # Apply language-specific processing
            if lang in ["tha", "tha+eng"]:
                # Thai-specific processing
                # 1. Noise removal
                gray = cv2.fastNlMeansDenoising(gray, None, 10, 7, 21)
                
                # 2. Improved adaptive thresholding
                binary = cv2.adaptiveThreshold(
                    gray, 255, cv2.ADAPTIVE_THRESH_GAUSSIAN_C,
                    cv2.THRESH_BINARY, 11, 2
                )
                
                # 3. Remove small noise
                kernel = np.ones((1, 1), np.uint8)
                binary = cv2.morphologyEx(binary, cv2.MORPH_OPEN, kernel)
                
                # 4. Tesseract configuration
                custom_config = r'--psm 6 --oem 1 -c preserve_interword_spaces=1'
            else:
                # English-specific processing
                # 1. Enhance contrast
                gray = cv2.equalizeHist(gray)
                
                # 2. Noise reduction
                gray = cv2.GaussianBlur(gray, (3, 3), 0)
                
                # 3. Otsu's thresholding for better binarization
                _, binary = cv2.threshold(gray, 0, 255, cv2.THRESH_BINARY + cv2.THRESH_OTSU)
                
                # 4. Deskew if needed
                binary = self.deskew(binary)
                
                # 5. Tesseract configuration
                custom_config = r'--psm 6 --oem 3'
            
            # Use pytesseract with enhanced configuration
            text = pytesseract.image_to_string(binary, lang=lang, config=custom_config)
            
            # Return text but preserve line breaks (only strip trailing/leading whitespace)
            return text.strip()
        
        except Exception as e:
            messagebox.showerror("Error", f"Failed to extract text: {e}")
            return ""
        
    def extract_and_display(self, image_num):
        """Extract text from the specified image and display it in the text area"""
        if image_num == 1:
            if not self.image1_path:
                messagebox.showwarning("Warning", "Please select image 1 first!")
                return
            language = self.lang1_var.get()
            extracted_text = self.extract_text(self.image1_path, language)
            self.image1_text = extracted_text
            self.text1.delete(1.0, tk.END)
            self.text1.insert(tk.END, extracted_text)
        else:
            if not self.image2_path:
                messagebox.showwarning("Warning", "Please select image 2 first!")
                return
            language = self.lang2_var.get()
            extracted_text = self.extract_text(self.image2_path, language)
            self.image2_text = extracted_text
            self.text2.delete(1.0, tk.END)
            self.text2.insert(tk.END, extracted_text)

    def deskew(self, image):
        """Deskew (straighten) text in the image"""
        try:
            # Find all white pixels
            coords = np.column_stack(np.where(image > 0))
            
            # Find the rotated rectangle
            angle = cv2.minAreaRect(coords)[-1]
            
            # Adjust angle
            if angle < -45:
                angle = -(90 + angle)
            else:
                angle = -angle
                
            # Rotate the image to deskew it if needed
            if abs(angle) > 0.5:  # Only rotate if angle is significant
                (h, w) = image.shape[:2]
                center = (w // 2, h // 2)
                M = cv2.getRotationMatrix2D(center, angle, 1.0)
                rotated = cv2.warpAffine(image, M, (w, h), flags=cv2.INTER_CUBIC, 
                                        borderMode=cv2.BORDER_REPLICATE)
                return rotated
                
            return image
        except:
            # If deskewing fails, return original image
            return image
    
    def compare_text(self):
        """Extract text from both images and compare them"""
        if not self.image1_path or not self.image2_path:
            messagebox.showwarning("Warning", "Please select both images first!")
            return
        
        try:
            # Extract text using selected language for each image
            self.image1_text = self.extract_text(self.image1_path, self.lang1_var.get())
            self.image2_text = self.extract_text(self.image2_path, self.lang2_var.get())
            
            # Clear previous text and tags
            self.text1.delete(1.0, tk.END)
            self.text2.delete(1.0, tk.END)
            
            # Calculate similarity
            if not self.image1_text and not self.image2_text:
                self.result_label.config(text="Both images have no detectable text.", fg="black")
                self.text1.insert(tk.END, "No text detected")
                self.text2.insert(tk.END, "No text detected")
                return
            
            # Split text by lines
            text1_lines = self.image1_text.splitlines()
            text2_lines = self.image2_text.splitlines()
            
            # Configure text tags for highlighting
            self.text1.tag_configure('diff', background='pink', foreground='red')
            self.text2.tag_configure('diff', background='pink', foreground='red')
            
            # Import difflib for word-by-word comparison
            import difflib
            
            # Track which lines have differences
            diff_lines = set()
            
            # Display each line with line numbers and highlight differences
            for i, (line1, line2) in enumerate(zip(text1_lines, text2_lines), 1):
                # Add line number prefix
                prefix1 = f"{i}: "
                prefix2 = f"{i}: "
                
                self.text1.insert(tk.END, prefix1)
                self.text2.insert(tk.END, prefix2)
                
                # Get current position after the line number
                line1_start = self.text1.index(tk.END + "-1c")
                line2_start = self.text2.index(tk.END + "-1c")
                
                # Split lines into words for better comparison
                words1 = line1.split()
                words2 = line2.split()
                
                if words1 != words2:
                    diff_lines.add(i)
                    
                    # Compare words
                    matcher = difflib.SequenceMatcher(None, words1, words2)
                    
                    # Insert text with appropriate tags
                    curr_pos1 = line1_start
                    curr_pos2 = line2_start
                    
                    for tag, i1, i2, j1, j2 in matcher.get_opcodes():
                        # Text in line 1
                        if i1 < i2:  # There are words to insert
                            if tag == 'equal':
                                self.text1.insert(tk.END, ' '.join(words1[i1:i2]) + ' ')
                            else:  # 'replace' or 'delete'
                                self.text1.insert(tk.END, ' '.join(words1[i1:i2]) + ' ', 'diff')
                        
                        # Text in line 2
                        if j1 < j2:  # There are words to insert
                            if tag == 'equal':
                                self.text2.insert(tk.END, ' '.join(words2[j1:j2]) + ' ')
                            else:  # 'replace' or 'insert'
                                self.text2.insert(tk.END, ' '.join(words2[j1:j2]) + ' ', 'diff')
                else:
                    # Lines are identical, just insert them
                    self.text1.insert(tk.END, line1)
                    self.text2.insert(tk.END, line2)
                
                # Add newline after each line
                self.text1.insert(tk.END, '\n')
                self.text2.insert(tk.END, '\n')
            
            # Handle extra lines in text1
            for i, line in enumerate(text1_lines[len(text2_lines):], len(text2_lines) + 1):
                diff_lines.add(i)
                prefix = f"{i}: "
                self.text1.insert(tk.END, prefix)
                self.text1.insert(tk.END, line, 'diff')
                self.text1.insert(tk.END, '\n')
            
            # Handle extra lines in text2
            for i, line in enumerate(text2_lines[len(text1_lines):], len(text1_lines) + 1):
                diff_lines.add(i)
                prefix = f"{i}: "
                self.text2.insert(tk.END, prefix)
                self.text2.insert(tk.END, line, 'diff')
                self.text2.insert(tk.END, '\n')
            
            # Calculate overall similarity
            text1_joined = ' '.join(text1_lines).lower()
            text2_joined = ' '.join(text2_lines).lower()
            distance = Levenshtein.distance(text1_joined, text2_joined)
            max_len = max(len(text1_joined), len(text2_joined))
            similarity = ((max_len - distance) / max_len) * 100 if max_len > 0 else 100
            
            # Always display similarity score regardless of result
            if len(text1_lines) != len(text2_lines):
                self.result_label.config(
                    text=f"✗ Different number of lines: Image 1 ({len(text1_lines)}) vs Image 2 ({len(text2_lines)}) (Similarity: {similarity:.1f}%)", 
                    fg="red"
                )
            elif diff_lines:
                self.result_label.config(
                    text=f"✗ Differences found in lines: {', '.join(map(str, sorted(diff_lines)))} (Similarity: {similarity:.1f}%)", 
                    fg="red"
                )
            elif similarity < 100:
                self.result_label.config(
                    text=f"✗ Differences in whitespace/formatting detected (Similarity: {similarity:.1f}%)", 
                    fg="orange"
                )
            else:
                self.result_label.config(
                    text=f"✓ The text in both images is the same. (Similarity: 100.0%)", 
                    fg="green"
                )
                
        except Exception as e:
            messagebox.showerror("Error", f"Comparison failed: {e}")
            import traceback
            traceback.print_exc()
    
    def show_visual_diff(self):
        """Generate and display a visual difference comparison"""
        if not self.image1_path or not self.image2_path or self.image1_cv is None or self.image2_cv is None:
            messagebox.showwarning("Warning", "Please select both images first!")
            return
            
        try:
            # Resize images to same dimensions for comparison
            # Get the dimensions of both images
            h1, w1 = self.image1_cv.shape[:2]
            h2, w2 = self.image2_cv.shape[:2]
            
            # Use the maximum dimensions for the comparison
            max_height = max(h1, h2)
            max_width = max(w1, w2)
            
            # Resize both images to the same dimensions
            img1_resized = cv2.resize(self.image1_cv, (max_width, max_height))
            img2_resized = cv2.resize(self.image2_cv, (max_width, max_height))
            
            # Convert to grayscale for better difference visualization
            gray1 = cv2.cvtColor(img1_resized, cv2.COLOR_BGR2GRAY)
            gray2 = cv2.cvtColor(img2_resized, cv2.COLOR_BGR2GRAY)
            
            # Calculate absolute difference between images
            diff = cv2.absdiff(gray1, gray2)
            
            # Create a color diff image - black background with red differences
            diff_color = np.zeros((max_height, max_width, 3), dtype=np.uint8)
            diff_color[diff > 30] = [0, 0, 255]  # Red where differences are significant
            
            # Create a side-by-side comparison
            img1_rgb = cv2.cvtColor(img1_resized, cv2.COLOR_BGR2RGB)
            img2_rgb = cv2.cvtColor(img2_resized, cv2.COLOR_BGR2RGB)
            
            # Create a 1x3 comparison image: img1, diff, img2
            comparison_width = max_width * 3
            comparison_img = np.zeros((max_height, comparison_width, 3), dtype=np.uint8)
            
            # Place images side by side
            comparison_img[:, :max_width] = img1_rgb
            comparison_img[:, max_width:max_width*2] = diff_color
            comparison_img[:, max_width*2:] = img2_rgb
            
            # Add text labels
            font = cv2.FONT_HERSHEY_SIMPLEX
            cv2.putText(comparison_img, 'Image 1', (10, 30), font, 1, (255, 255, 255), 2)
            cv2.putText(comparison_img, 'Differences (Red)', (max_width + 10, 30), font, 1, (255, 255, 255), 2)
            cv2.putText(comparison_img, 'Image 2', (max_width*2 + 10, 30), font, 1, (255, 255, 255), 2)
            
            # Convert to PIL format for ZoomableCanvas
            comparison_pil = Image.fromarray(comparison_img)
            
            # Set the image to our zoomable canvas
            self.visual_canvas.set_image(comparison_pil)
            
            # Switch to the visual comparison tab
            self.notebook.select(self.visual_tab)
            
        except Exception as e:
            messagebox.showerror("Error", f"Failed to generate visual difference: {e}")
            import traceback
            traceback.print_exc()
    
    def reset(self):
        """Reset the application to its initial state"""
        # Clear images
        self.image1_path = None
        self.image2_path = None
        self.image1_cv = None
        self.image2_cv = None
        self.image1_label.config(image="", text="No image selected")
        self.image2_label.config(image="", text="No image selected")
        
        # Reset language selections to English
        self.lang1_var.set("eng")
        self.lang2_var.set("eng")
        
        # Clear text areas
        self.text1.delete(1.0, tk.END)
        self.text2.delete(1.0, tk.END)
        
        # Reset result
        self.result_label.config(text="No comparison performed yet", fg="black")
        
        # Reset visual canvas
        self.visual_canvas.set_text("No visual comparison available\nClick 'Show Visual Diff' button to generate")
        
        # Clear stored text
        self.image1_text = ""
        self.image2_text = ""
        
        # Switch to text tab
        self.notebook.select(self.text_tab)
        
        # Also reset detection tab
        self.detection_canvas.set_text("No detection analysis performed\nClick 'Detect Differences' to analyze images")
        self.detection_details.delete(1.0, tk.END)
        self.sensitivity_var.set(50.0)  # Reset sensitivity to default
        self.method_var.set("contour")  # Reset detection method to default
        
        # Also reset bill verification tab
        if hasattr(self, 'bill_path'):
            self.reset_bill_verification()

    def detect_counterfeit(self):
        """Detect and highlight differences between images using the selected method"""
        if not self.image1_path or not self.image2_path or self.image1_cv is None or self.image2_cv is None:
            messagebox.showwarning("Warning", "Please select both images first!")
            return
            
        try:
            # Get detection sensitivity (1-100)
            sensitivity = self.sensitivity_var.get()
            # Normalize to 0-1 range for calculations
            sensitivity_norm = sensitivity / 100.0
            
            # Get detection method
            method = self.method_var.get()
            
            # Clear previous detection details
            self.detection_details.delete(1.0, tk.END)
            
            # Resize images to same dimensions for comparison
            h1, w1 = self.image1_cv.shape[:2]
            h2, w2 = self.image2_cv.shape[:2]
            
            # Use the maximum dimensions for the comparison
            max_height = max(h1, h2)
            max_width = max(w1, w2)
            
            # Resize both images to the same dimensions
            img1_resized = cv2.resize(self.image1_cv, (max_width, max_height))
            img2_resized = cv2.resize(self.image2_cv, (max_width, max_height))
            
            # Create a side-by-side comparison
            img1_rgb = cv2.cvtColor(img1_resized, cv2.COLOR_BGR2RGB)
            img2_rgb = cv2.cvtColor(img2_resized, cv2.COLOR_BGR2RGB)
            
            # Create a combined image: img1 | img2
            combined_width = max_width * 2
            result_img = np.zeros((max_height, combined_width, 3), dtype=np.uint8)
            
            # Place images side by side
            result_img[:, :max_width] = img1_rgb
            result_img[:, max_width:] = img2_rgb
            
            # Detected regions (will be populated based on method)
            regions_img1 = []
            regions_img2 = []
            
            # Use different detection methods based on user selection
            if method == "contour":
                # Contour-based detection
                regions_img1, regions_img2 = self.detect_contour_differences(
                    img1_resized, img2_resized, sensitivity_norm)
            elif method == "feature":
                # Feature matching detection
                regions_img1, regions_img2 = self.detect_feature_differences(
                    img1_resized, img2_resized, sensitivity_norm)
            
            # Draw bounding boxes on the combined image
            for (x, y, w, h) in regions_img1:
                cv2.rectangle(result_img, (x, y), (x + w, y + h), (255, 0, 0), 2)
                # Add label
                cv2.putText(result_img, "Suspicious", (x, y-5), cv2.FONT_HERSHEY_SIMPLEX, 0.5, (255, 0, 0), 1)
                
            for (x, y, w, h) in regions_img2:
                # Adjust x-coordinate for second image
                adjusted_x = x + max_width
                cv2.rectangle(result_img, (adjusted_x, y), (adjusted_x + w, y + h), (255, 0, 0), 2)
                # Add label
                cv2.putText(result_img, "Suspicious", (adjusted_x, y-5), cv2.FONT_HERSHEY_SIMPLEX, 0.5, (255, 0, 0), 1)
                
            # Add text labels for images
            font = cv2.FONT_HERSHEY_SIMPLEX
            cv2.putText(result_img, 'Original', (10, 30), font, 1, (255, 255, 255), 2)
            cv2.putText(result_img, 'Comparison', (max_width + 10, 30), font, 1, (255, 255, 255), 2)
            
            # Convert to PIL format for display
            detection_pil = Image.fromarray(result_img)
            
            # Set the image to our zoomable canvas
            self.detection_canvas.set_image(detection_pil)
            
            # Show detection summary
            total_regions = len(regions_img1) + len(regions_img2)
            confidence_score = self.calculate_confidence_score(regions_img1, regions_img2, sensitivity_norm)
            
            self.detection_details.insert(tk.END, f"Detection Method: {method.capitalize()}\n")
            self.detection_details.insert(tk.END, f"Sensitivity: {int(sensitivity)}%\n")
            self.detection_details.insert(tk.END, f"Suspicious Areas Detected: {total_regions}\n")
            self.detection_details.insert(tk.END, f"Confidence Score: {confidence_score:.1f}%\n\n")
            
            if total_regions > 0:
                self.detection_details.insert(tk.END, f"Detected {len(regions_img1)} suspicious areas in image 1\n")
                self.detection_details.insert(tk.END, f"Detected {len(regions_img2)} suspicious areas in image 2\n\n")
                
                if confidence_score > 75:
                    assessment = "HIGH RISK: Significant manipulations detected!"
                    self.detection_details.insert(tk.END, assessment, "high_risk")
                elif confidence_score > 30:
                    assessment = "MEDIUM RISK: Some suspicious differences detected."
                    self.detection_details.insert(tk.END, assessment, "medium_risk")
                else:
                    assessment = "LOW RISK: Minor differences detected."
                    self.detection_details.insert(tk.END, assessment, "low_risk")
            else:
                self.detection_details.insert(tk.END, "No suspicious areas detected. Images appear legitimate.")
            
            # Configure text tags for risk levels
            self.detection_details.tag_configure('high_risk', background='#ffcccc', foreground='#cc0000')
            self.detection_details.tag_configure('medium_risk', background='#ffffcc', foreground='#cc6600')
            self.detection_details.tag_configure('low_risk', background='#e6ffcc', foreground='#006600')
            
            # Switch to the detection tab
            self.notebook.select(self.detection_tab)
            
        except Exception as e:
            messagebox.showerror("Error", f"Detection failed: {e}")
            import traceback
            traceback.print_exc()

    def detect_contour_differences(self, img1, img2, sensitivity):
        """Detect differences using contour detection"""
        # Convert to grayscale
        gray1 = cv2.cvtColor(img1, cv2.COLOR_BGR2GRAY)
        gray2 = cv2.cvtColor(img2, cv2.COLOR_BGR2GRAY)
        
        # Apply Gaussian blur to reduce noise
        gray1 = cv2.GaussianBlur(gray1, (5, 5), 0)
        gray2 = cv2.GaussianBlur(gray2, (5, 5), 0)
        
        # Calculate absolute difference between images
        diff = cv2.absdiff(gray1, gray2)
        
        # Apply threshold - adjust based on sensitivity
        threshold_value = int(255 * (1 - sensitivity))  # Higher sensitivity = lower threshold
        threshold_value = max(10, min(threshold_value, 240))  # Keep within reasonable bounds
        
        _, thresh = cv2.threshold(diff, threshold_value, 255, cv2.THRESH_BINARY)
        
        # Dilate the threshold image to fill in holes
        kernel = np.ones((5, 5), np.uint8)
        dilated = cv2.dilate(thresh, kernel, iterations=3)
        
        # Find contours
        contours, _ = cv2.findContours(dilated, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)
        
        # Filter contours by area to avoid noise
        min_area = 50  # Minimum area to consider (can be adjusted)
        valid_contours = [cnt for cnt in contours if cv2.contourArea(cnt) > min_area]
        
        # Get bounding rectangles for both images (all on first image for contour method)
        regions_img1 = [cv2.boundingRect(cnt) for cnt in valid_contours]
        regions_img2 = []  # For contour method, we allocate all regions to img1
        
        return regions_img1, regions_img2

    def detect_feature_differences(self, img1, img2, sensitivity):
        """Detect differences using feature matching"""
        # Convert to grayscale
        gray1 = cv2.cvtColor(img1, cv2.COLOR_BGR2GRAY)
        gray2 = cv2.cvtColor(img2, cv2.COLOR_BGR2GRAY)
        
        # Initialize feature detector (ORB is generally available without additional installation)
        max_features = int(10000 * sensitivity)  # More features with higher sensitivity
        feature_detector = cv2.ORB_create(nfeatures=max_features)
        
        # Find keypoints and descriptors
        keypoints1, descriptors1 = feature_detector.detectAndCompute(gray1, None)
        keypoints2, descriptors2 = feature_detector.detectAndCompute(gray2, None)
        
        # If no features found in either image
        if descriptors1 is None or descriptors2 is None or len(keypoints1) == 0 or len(keypoints2) == 0:
            return [], []
        
        # Create feature matcher
        bf = cv2.BFMatcher(cv2.NORM_HAMMING, crossCheck=True)
        
        # Match descriptors
        matches = bf.match(descriptors1, descriptors2)
        
        # Sort matches by distance (lower distance = better match)
        matches = sorted(matches, key=lambda x: x.distance)
        
        # Keep only good matches based on sensitivity
        # Higher sensitivity means fewer matches considered "good" (more suspicious areas)
        good_match_percent = 1.0 - sensitivity
        num_good_matches = int(len(matches) * good_match_percent)
        good_matches = matches[:num_good_matches]
        
        # Find points that didn't match well
        matched_keypoints1 = set([good_matches[i].queryIdx for i in range(len(good_matches))])
        matched_keypoints2 = set([good_matches[i].trainIdx for i in range(len(good_matches))])
        
        # Find unmatched keypoints (suspicious areas)
        unmatched_keypoints1 = [keypoints1[i] for i in range(len(keypoints1)) if i not in matched_keypoints1]
        unmatched_keypoints2 = [keypoints2[i] for i in range(len(keypoints2)) if i not in matched_keypoints2]
        
        # Group nearby keypoints into regions
        regions_img1 = self.group_keypoints_into_regions(unmatched_keypoints1, gray1.shape)
        regions_img2 = self.group_keypoints_into_regions(unmatched_keypoints2, gray2.shape)
        
        return regions_img1, regions_img2

    def group_keypoints_into_regions(self, keypoints, img_shape, distance_threshold=50):
        """Group nearby keypoints into bounding box regions"""
        if not keypoints:
            return []
            
        # Extract points
        points = np.array([(int(kp.pt[0]), int(kp.pt[1])) for kp in keypoints])
        
        # If we have too few points, just return a bounding box around them all
        if len(points) < 5:
            if len(points) == 0:
                return []
                
            x_min = np.min(points[:, 0])
            y_min = np.min(points[:, 1])
            x_max = np.max(points[:, 0])
            y_max = np.max(points[:, 1])
            
            width = x_max - x_min + 20  # Add padding
            height = y_max - y_min + 20
            
            return [(max(0, x_min - 10), max(0, y_min - 10), width, height)]
        
        # Use clustering to group nearby points
        # We'll use DBSCAN for clustering points based on proximity
        from sklearn.cluster import DBSCAN
        
        clustering = DBSCAN(eps=distance_threshold, min_samples=2).fit(points)
        labels = clustering.labels_
        
        # Extract regions from clusters
        regions = []
        unique_labels = set(labels)
        
        for label in unique_labels:
            # Skip noise points (label -1)
            if label == -1:
                continue
                
            # Get points in this cluster
            cluster_points = points[labels == label]
            
            # Calculate bounding box
            x_min = np.min(cluster_points[:, 0])
            y_min = np.min(cluster_points[:, 1])
            x_max = np.max(cluster_points[:, 0])
            y_max = np.max(cluster_points[:, 1])
            
            width = x_max - x_min + 20  # Add padding
            height = y_max - y_min + 20
            
            regions.append((max(0, x_min - 10), max(0, y_min - 10), width, height))
        
        # Add isolated points (noise) as small regions
        for i, point in enumerate(points):
            if labels[i] == -1:
                x, y = point
                regions.append((max(0, x - 10), max(0, y - 10), 20, 20))
        
        return regions

    def calculate_confidence_score(self, regions_img1, regions_img2, sensitivity):
        """Calculate a confidence score for counterfeit detection"""
        # Count total suspicious regions
        total_regions = len(regions_img1) + len(regions_img2)
        
        # Calculate total area of suspicious regions
        total_area_img1 = sum([w * h for _, _, w, h in regions_img1])
        total_area_img2 = sum([w * h for _, _, w, h in regions_img2])
        
        # Get image dimensions (assuming both are same size now)
        if self.image1_cv is not None and self.image2_cv is not None:
            h1, w1 = self.image1_cv.shape[:2]
            total_image_area = h1 * w1
            
            # Calculate percentage of suspicious area
            suspicious_area_percent = (total_area_img1 + total_area_img2) / (total_image_area * 2) * 100
        else:
            suspicious_area_percent = 0
        
        # Combine factors for confidence score
        # More regions, more area, higher sensitivity = higher confidence score
        base_score = min(100, suspicious_area_percent * 2)
        
        # Adjust based on regions count
        region_factor = min(100, total_regions * 10)
        
        # Adjust based on sensitivity
        sensitivity_factor = sensitivity * 100
        
        # Weighted combination
        confidence_score = (base_score * 0.5) + (region_factor * 0.3) + (sensitivity_factor * 0.2)
        
        # Cap at 100%
        return min(100, confidence_score)

    def browse_bill(self):
        """Open a file dialog to select a bill/receipt to verify"""
        file_path = filedialog.askopenfilename(
            title="Select Document to Verify",
            filetypes=(("Image files", "*.jpg;*.jpeg;*.png;*.bmp"), ("All files", "*.*"))
        )
        
        if file_path:
            try:
                # Store the original image for processing
                cv_img = cv2.imdecode(np.fromfile(file_path, dtype=np.uint8), cv2.IMREAD_UNCHANGED)
                
                # Load and resize image for display
                img = Image.open(file_path)
                
                # Store bill image info
                self.bill_path = file_path
                self.bill_cv = cv_img
                
                # Display in canvas
                self.bill_canvas.set_image(img)
                
                # Reset fields
                self.tx_id_var.set("")
                self.amount_var.set("")
                self.datetime_var.set("")
                
                # Reset results
                self.results_text.config(state=tk.NORMAL)
                self.results_text.delete(1.0, tk.END)
                self.results_text.insert(tk.END, "Document loaded. Use verification tools to analyze.")
                self.results_text.config(state=tk.DISABLED)
                
            except Exception as e:
                messagebox.showerror("Error", f"Failed to open document: {e}")

    def auto_extract(self, field_type):
        """Automatically extract specific field data from the bill"""
        if self.bill_cv is None:
            messagebox.showwarning("Warning", "Please upload a document first!")
            return
        
        try:
            # Extract text from the bill
            doc_text = self.extract_text(self.bill_path, "eng+tha")
            
            # Based on document type and field, look for patterns
            if field_type == "tx_id":
                # Extract transaction ID based on document type
                if self.bill_type_var.get() == "banking":
                    # Look for patterns like "Ref: XXXNNNNNNNNN" or "Transaction: XXXNNNNNNNNN"
                    tx_patterns = [
                        r"Ref[.:#]\s*([A-Z0-9]{8,15})",
                        r"Transaction[.:#]\s*([A-Z0-9]{8,15})",
                        r"เลขที่อ้างอิง[.:#]\s*([A-Z0-9]{8,15})"
                    ]
                    
                    for pattern in tx_patterns:
                        match = re.search(pattern, doc_text)
                        if match:
                            self.tx_id_var.set(match.group(1))
                            return
                    
                    # If no match, try to find any sequence that looks like a transaction ID
                    match = re.search(r"[A-Z]{2,3}[0-9]{6,10}", doc_text)
                    if match:
                        self.tx_id_var.set(match.group(0))
                        return
                    
                elif self.bill_type_var.get() == "receipt":
                    # Look for receipt number patterns
                    receipt_patterns = [
                        r"Receipt No[.:#]\s*([A-Z0-9]{4,15})",
                        r"ใบเสร็จรับเงินเลขที่[.:#]\s*([A-Z0-9]{4,15})"
                    ]
                    
                    for pattern in receipt_patterns:
                        match = re.search(pattern, doc_text)
                        if match:
                            self.tx_id_var.set(match.group(1))
                            return
                            
                # QR payment - try generic approach
                match = re.search(r"[A-Z0-9]{12,20}", doc_text)
                if match:
                    self.tx_id_var.set(match.group(0))
                    return
                    
                messagebox.showinfo("Info", "Couldn't automatically detect Transaction ID. Please enter manually.")
                    
            elif field_type == "amount":
                # Try to find amount patterns
                amount_patterns = [
                    r"Amount[.:#]?\s*(?:THB|฿)?\s*([0-9,.]+)",
                    r"จำนวนเงิน[.:#]?\s*(?:บาท)?\s*([0-9,.]+)",
                    r"(?:THB|฿)\s*([0-9,.]+)",
                    r"([0-9,.]+)\s*(?:บาท|THB|฿)"
                ]
                
                for pattern in amount_patterns:
                    match = re.search(pattern, doc_text)
                    if match:
                        amount = match.group(1).replace(',', '')
                        try:
                            # Validate it's really a number
                            float(amount)
                            self.amount_var.set(amount)
                            return
                        except:
                            continue
                
                messagebox.showinfo("Info", "Couldn't automatically detect Amount. Please enter manually.")
                    
            elif field_type == "datetime":
                # Try to find date patterns
                date_patterns = [
                    r"Date[.:#]?\s*([0-9]{1,2}[/-][0-9]{1,2}[/-][0-9]{2,4})",
                    r"Date[.:#]?\s*([A-Za-z]{3,9}\s*[0-9]{1,2},?\s*[0-9]{2,4})",
                    r"วันที่[.:#]?\s*([0-9]{1,2}[/-][0-9]{1,2}[/-][0-9]{2,4})",
                    r"([0-9]{1,2}[/-][0-9]{1,2}[/-][0-9]{2,4})\s*(?:เวลา)?\s*([0-9]{1,2}:[0-9]{2})"
                ]
                
                for pattern in date_patterns:
                    match = re.search(pattern, doc_text)
                    if match:
                        date_str = match.group(1)
                        self.datetime_var.set(date_str)
                        return
                
                messagebox.showinfo("Info", "Couldn't automatically detect Date/Time. Please enter manually.")
                
        except Exception as e:
            messagebox.showerror("Error", f"Extraction failed: {e}")

    def scan_qr_code(self):
        """Scan and validate QR code in the document"""
        if self.bill_cv is None:
            messagebox.showwarning("Warning", "Please upload a document first!")
            return
            
        try:
            # Convert to grayscale
            gray = cv2.cvtColor(self.bill_cv, cv2.COLOR_BGR2GRAY)
            
            # Use cv2's QR code detector
            qr_detector = cv2.QRCodeDetector()
            data, bbox, _ = qr_detector.detectAndDecode(gray)
            
            if data:
                # Update results text
                self.results_text.config(state=tk.NORMAL)
                self.results_text.delete(1.0, tk.END)
                self.results_text.insert(tk.END, f"QR Code Detected!\n\nEncoded Data:\n{data}\n\n")
                
                # Analyze QR data based on document type
                if self.bill_type_var.get() == "banking":
                    self.validate_banking_qr(data)
                elif self.bill_type_var.get() == "receipt":
                    self.validate_receipt_qr(data)
                else:
                    self.validate_payment_qr(data)
                    
                # Draw bounding box around QR code if detected
                if bbox is not None and len(bbox) > 0:
                    # Make a copy of the original image
                    qr_highlight = self.bill_cv.copy()
                    
                    # Draw polygon around QR code
                    bbox = bbox.astype(int)
                    cv2.polylines(qr_highlight, [bbox], True, (0, 255, 0), 2)
                    
                    # Convert to PIL format for display
                    qr_highlight_rgb = cv2.cvtColor(qr_highlight, cv2.COLOR_BGR2RGB)
                    qr_pil = Image.fromarray(qr_highlight_rgb)
                    
                    # Update display
                    self.bill_canvas.set_image(qr_pil)
                
                self.results_text.config(state=tk.DISABLED)
            else:
                messagebox.showinfo("Result", "No QR code found in the image.")
        except Exception as e:
            messagebox.showerror("Error", f"QR code scanning failed: {e}")

    def validate_banking_qr(self, qr_data):
        """Validate QR data for banking slips"""
        try:
            # Check for common banking QR data patterns
            bill_pay_pattern = r"billpay|payment|bank"
            transfer_pattern = r"transfer|โอนเงิน"
            amount_pattern = r"amount=([0-9.]+)"
            
            # Simple validation check
            validation_score = 0
            
            if re.search(bill_pay_pattern, qr_data, re.IGNORECASE):
                validation_score += 30
                self.results_text.insert(tk.END, "✓ QR contains valid banking payment data\n")
            
            if re.search(transfer_pattern, qr_data, re.IGNORECASE):
                validation_score += 30
                self.results_text.insert(tk.END, "✓ QR contains valid transfer information\n")
                
            amount_match = re.search(amount_pattern, qr_data, re.IGNORECASE)
            if amount_match:
                # Extract amount from QR
                qr_amount = amount_match.group(1)
                
                # Compare with user-entered amount if available
                if self.amount_var.get():
                    try:
                        user_amount = self.amount_var.get().replace(',', '')
                        if float(user_amount) == float(qr_amount):
                            validation_score += 40
                            self.results_text.insert(tk.END, f"✓ Amount in QR ({qr_amount}) matches entered amount\n")
                        else:
                            self.results_text.insert(tk.END, f"✗ Amount mismatch: QR={qr_amount}, Entered={user_amount}\n")
                    except:
                        pass
                else:
                    # Automatically set the amount
                    self.amount_var.set(qr_amount)
                    validation_score += 20
            
            # Final assessment
            self.results_text.insert(tk.END, f"\nQR Validation Score: {validation_score}%\n\n")
            
            if validation_score >= 60:
                self.results_text.insert(tk.END, "QR CODE VERIFIED: Contains valid banking information")
            else:
                self.results_text.insert(tk.END, "SUSPICIOUS QR CODE: Missing expected banking data patterns")
                
        except Exception as e:
            self.results_text.insert(tk.END, f"Error validating QR: {e}")

    def validate_receipt_qr(self, qr_data):
        """Validate QR data for payment receipts"""
        try:
            # Check for common receipt QR data patterns
            receipt_pattern = r"receipt|invoice|ใบเสร็จ|ใบสำคัญรับเงิน"
            vendor_pattern = r"vendor=|merchant=|shop="
            tax_pattern = r"tax|vat|ภาษี"
            
            # Simple validation check
            validation_score = 0
            
            if re.search(receipt_pattern, qr_data, re.IGNORECASE):
                validation_score += 30
                self.results_text.insert(tk.END, "✓ QR contains valid receipt information\n")
            
            if re.search(vendor_pattern, qr_data, re.IGNORECASE):
                validation_score += 20
                self.results_text.insert(tk.END, "✓ QR contains vendor/merchant information\n")
                
            if re.search(tax_pattern, qr_data, re.IGNORECASE):
                validation_score += 20
                self.results_text.insert(tk.END, "✓ QR contains tax/VAT information\n")
                
            # Look for date pattern
            date_match = re.search(r"date=(\S+)", qr_data, re.IGNORECASE)
            if date_match:
                validation_score += 30
                date_str = date_match.group(1)
                self.results_text.insert(tk.END, f"✓ QR contains date information: {date_str}\n")
                
                # Set date if not already set
                if not self.datetime_var.get():
                    self.datetime_var.set(date_str)
            
            # Final assessment
            self.results_text.insert(tk.END, f"\nQR Validation Score: {validation_score}%\n\n")
            
            if validation_score >= 50:
                self.results_text.insert(tk.END, "QR CODE VERIFIED: Contains valid receipt information")
            else:
                self.results_text.insert(tk.END, "SUSPICIOUS QR CODE: Missing expected receipt data patterns")
                
        except Exception as e:
            self.results_text.insert(tk.END, f"Error validating QR: {e}")

    def validate_payment_qr(self, qr_data):
        """Validate QR data for QR payments (PromptPay/QR PromptPay)"""
        try:
            # Check for common QR payment patterns
            promptpay_pattern = r"promptpay|พร้อมเพย์"
            amount_pattern = r"amount=([0-9.]+)"
            merchant_pattern = r"merchant=|receiver="
            
            # Simple validation check
            validation_score = 0
            
            if re.search(promptpay_pattern, qr_data, re.IGNORECASE):
                validation_score += 40
                self.results_text.insert(tk.END, "✓ QR contains valid PromptPay data\n")
            
            merchant_match = re.search(merchant_pattern, qr_data, re.IGNORECASE)
            if merchant_match:
                validation_score += 30
                self.results_text.insert(tk.END, "✓ QR contains merchant/recipient information\n")
                
            amount_match = re.search(amount_pattern, qr_data, re.IGNORECASE)
            if amount_match:
                # Extract amount from QR
                qr_amount = amount_match.group(1)
                validation_score += 30
                
                # Compare with user-entered amount if available
                if self.amount_var.get():
                    try:
                        user_amount = self.amount_var.get().replace(',', '')
                        if float(user_amount) == float(qr_amount):
                            self.results_text.insert(tk.END, f"✓ Amount in QR ({qr_amount}) matches entered amount\n")
                        else:
                            validation_score -= 20
                            self.results_text.insert(tk.END, f"✗ Amount mismatch: QR={qr_amount}, Entered={user_amount}\n")
                    except:
                        pass
                else:
                    # Automatically set the amount
                    self.amount_var.set(qr_amount)
                    self.results_text.insert(tk.END, f"✓ Amount detected from QR: {qr_amount}\n")
            
            # Final assessment
            self.results_text.insert(tk.END, f"\nQR Validation Score: {validation_score}%\n\n")
            
            if validation_score >= 60:
                self.results_text.insert(tk.END, "QR CODE VERIFIED: Contains valid payment information")
            else:
                self.results_text.insert(tk.END, "SUSPICIOUS QR CODE: Missing expected payment data patterns")
                
        except Exception as e:
            self.results_text.insert(tk.END, f"Error validating QR: {e}")

    def verify_bill(self):
        """Verify the uploaded document for authenticity"""
        if self.bill_cv is None:
            messagebox.showwarning("Warning", "Please upload a document first!")
            return
        
        try:
            # Enable text widget for updating
            self.results_text.config(state=tk.NORMAL)
            self.results_text.delete(1.0, tk.END)
            
            # Initialize verification score
            verification_score = 0
            max_score = 100
            checks_passed = 0
            checks_total = 5
            verification_notes = []
            
            # 1. Extract text from the image for content verification
            doc_text = self.extract_text(self.bill_path, "eng+tha")
            
            # 2. Check for transaction ID format
            if self.tx_id_var.get():
                tx_id = self.tx_id_var.get()
                if self.validate_transaction_id(tx_id):
                    verification_score += 20
                    checks_passed += 1
                    verification_notes.append("✓ Transaction ID format is valid")
                else:
                    verification_notes.append("✗ Transaction ID format is invalid or suspicious")
            
            # 3. Check for date/time consistency
            if self.datetime_var.get():
                datetime_str = self.datetime_var.get()
                if self.validate_datetime(datetime_str):
                    verification_score += 20
                    checks_passed += 1
                    verification_notes.append("✓ Date/time format is valid")
                else:
                    verification_notes.append("✗ Date/time is invalid or inconsistent")
            
            # 4. Check image quality and artifacts
            image_score = self.analyze_image_quality(self.bill_cv)
            verification_score += image_score
            if image_score > 15:
                checks_passed += 1
                verification_notes.append("✓ Image quality consistent with genuine document")
            else:
                verification_notes.append("✗ Image shows signs of manipulation or generation")
            
            # 5. Analyze document structure based on type
            if self.bill_type_var.get() == "banking":
                structure_valid = self.validate_banking_slip_structure(self.bill_cv, doc_text)
            elif self.bill_type_var.get() == "receipt":
                structure_valid = self.validate_receipt_structure(self.bill_cv, doc_text)
            else:
                structure_valid = self.validate_qr_payment_structure(self.bill_cv, doc_text)
                
            if structure_valid:
                verification_score += 20
                checks_passed += 1
                verification_notes.append("✓ Document structure matches expected template")
            else:
                verification_notes.append("✗ Document structure deviates from expected template")
            
            # 6. Content consistency (e.g., amount, calculations)
            if self.amount_var.get():
                amount = self.amount_var.get()
                if self.validate_amount_consistency(amount, doc_text):
                    verification_score += 20
                    checks_passed += 1
                    verification_notes.append("✓ Amount is consistently displayed throughout document")
                else:
                    verification_notes.append("✗ Amount inconsistencies detected")
            
            # Display verification result
            self.results_text.insert(tk.END, f"Verification Complete\n\n")
            self.results_text.insert(tk.END, f"Score: {verification_score}/{max_score}\n")
            self.results_text.insert(tk.END, f"Checks Passed: {checks_passed}/{checks_total}\n\n")
            
            # Apply appropriate tag based on verification score
            if verification_score >= 80:
                assessment = "HIGH CONFIDENCE: Document appears to be authentic."
                self.results_text.insert(tk.END, assessment + "\n\n", "authentic")
            elif verification_score >= 50:
                assessment = "MEDIUM CONFIDENCE: Document has some suspicious elements."
                self.results_text.insert(tk.END, assessment + "\n\n", "uncertain")
            else:
                assessment = "LOW CONFIDENCE: Document is likely fake or manipulated."
                self.results_text.insert(tk.END, assessment + "\n\n", "fake")
            
            # Display verification notes
            self.results_text.insert(tk.END, "Verification Notes:\n")
            for note in verification_notes:
                self.results_text.insert(tk.END, note + "\n")
            
            # Configure text tags for verification results
            self.results_text.tag_configure('authentic', background='#e6ffcc', foreground='#006600')
            self.results_text.tag_configure('uncertain', background='#ffffcc', foreground='#cc6600')
            self.results_text.tag_configure('fake', background='#ffcccc', foreground='#cc0000')
            
            self.results_text.config(state=tk.DISABLED)
            
        except Exception as e:
            messagebox.showerror("Error", f"Verification failed: {e}")

    def analyze_image_quality(self, image):
        """Analyze image for signs of AI generation or manipulation"""
        # Noise pattern analysis
        gray = cv2.cvtColor(image, cv2.COLOR_BGR2GRAY)
        
        # Look for abnormal noise patterns
        noise_level = np.std(gray)
        
        # Calculate image entropy (AI generated images often have lower entropy)
        hist = cv2.calcHist([gray], [0], None, [256], [0, 256])
        hist = hist / hist.sum()
        entropy = -np.sum(hist * np.log2(hist + 1e-7))
        
        # Edge coherence - AI images may have unnaturally smooth edges
        edges = cv2.Canny(gray, 100, 200)
        edge_coherence = np.sum(edges) / (gray.shape[0] * gray.shape[1])
        
        # JPEG artifacts analysis
        _, enc_img = cv2.imencode('.jpg', image, [cv2.IMWRITE_JPEG_QUALITY, 90])
        dec_img = cv2.imdecode(enc_img, cv2.IMREAD_COLOR)
        diff = cv2.absdiff(image, dec_img)
        jpeg_artifacts = np.mean(diff)
        
        # Score calculation (max 20 points)
        score = 0
        
        # Natural images typically have moderate noise
        if 5 < noise_level < 30:
            score += 5
        
        # Natural images have higher entropy
        if entropy > 7.0:
            score += 5
        
        # Natural images have natural edge coherence
        if 0.05 < edge_coherence < 0.2:
            score += 5
        
        # JPEG artifacts should be within normal range for real photos
        if 0.5 < jpeg_artifacts < 5:
            score += 5
        
        return score

    def validate_transaction_id(self, tx_id):
        """Validate transaction ID format"""
        # Check for common transaction ID patterns based on document type
        if self.bill_type_var.get() == "banking":
            # Banking slip transaction IDs often have patterns like XXXNNNNNNNNN
            return bool(re.match(r'^[A-Z]{2,3}\d{6,10}$', tx_id))
        elif self.bill_type_var.get() == "receipt":
            # Receipt reference numbers often follow patterns like YYYYMMDDNNNNN
            return bool(re.match(r'^\d{8,14}$', tx_id) or re.match(r'^[A-Z]{1,2}\d{6,10}$', tx_id))
        else:
            # QR payment references often have alphanumeric codes
            return bool(re.match(r'^[A-Za-z0-9]{8,16}$', tx_id))

    def validate_datetime(self, date_str):
        """Validate date format and check if it's reasonable"""
        # Try multiple formats
        formats = [
            "%d/%m/%Y", "%d-%m-%Y", "%Y/%m/%d", "%Y-%m-%d",
            "%d/%m/%y", "%d-%m-%y",
            "%b %d, %Y", "%B %d, %Y",
            "%d %b %Y", "%d %B %Y"
        ]
        
        # Try parsing with each format
        for fmt in formats:
            try:
                dt = datetime.strptime(date_str, fmt)
                # Check if date is reasonable (not in the future, not too old)
                now = datetime.now()
                if dt <= now and dt.year >= 2000:
                    return True
            except:
                continue
        
        # Thai date format check (e.g., "25 ก.พ. 2566")
        thai_month_pattern = r"(\d{1,2})\s+([กขคงจฉชซฌญฎฏฐฑฒณดตถทธนบปผฝพฟภมยรลวศษสหฬอฮ]{1,3}\.?)\s+(\d{4})"
        match = re.match(thai_month_pattern, date_str)
        if match:
            try:
                # Convert Thai year (Buddhist era) to Gregorian calendar if needed
                year = int(match.group(3))
                if year > 2500:  # Likely Buddhist Era
                    year -= 543  # Convert to CE/AD
                    
                if 2000 <= year <= datetime.now().year:
                    return True
            except:
                pass
        
        return False

    def validate_banking_slip_structure(self, image, text):
        """Validate the structure of a banking slip"""
        # Basic check for common banking slip elements
        banking_terms = [
            "transfer", "transaction", "reference", "amount", "date", "receipt",
            "bank", "payment", "account", "branch", "โอนเงิน", "ธนาคาร", "สาขา",
            "รายการ", "จำนวน", "บัญชี"
        ]
        
        count = 0
        for term in banking_terms:
            if term.lower() in text.lower():
                count += 1
        
        # Need at least 3 banking terms
        return count >= 3

    def validate_receipt_structure(self, image, text):
        """Validate the structure of a payment receipt"""
        # Basic check for common receipt elements
        receipt_terms = [
            "receipt", "invoice", "total", "amount", "payment", "date", "customer",
            "item", "price", "quantity", "tax", "vat", "ใบเสร็จ", "ใบกำกับภาษี",
            "รวมเงิน", "จำนวนเงิน", "ราคา", "ภาษีมูลค่าเพิ่ม"
        ]
        
        count = 0
        for term in receipt_terms:
            if term.lower() in text.lower():
                count += 1
        
        # Need at least 3 receipt terms
        return count >= 3

    def validate_qr_payment_structure(self, image, text):
        """Validate the structure of a QR payment"""
        # Basic check for common QR payment elements
        qr_terms = [
            "promptpay", "qr code", "payment", "amount", "ref", "scan",
            "พร้อมเพย์", "คิวอาร์", "สแกน", "ชำระเงิน", "จำนวนเงิน"
        ]
        
        count = 0
        for term in qr_terms:
            if term.lower() in text.lower():
                count += 1
        
        # Check for presence of QR code
        gray = cv2.cvtColor(image, cv2.COLOR_BGR2GRAY)
        qr_detector = cv2.QRCodeDetector()
        data, bbox, _ = qr_detector.detectAndDecode(gray)
        
        if bbox is not None and len(bbox) > 0:
            count += 3  # Strong indicator - actual QR code present
        
        # Need at least 3 QR terms or presence of actual QR code
        return count >= 3

    def validate_amount_consistency(self, amount_str, text):
        """Check if amount appears consistently in the document"""
        try:
            # Clean and format amount string for comparison
            amount = amount_str.replace(',', '')
            amount_float = float(amount)
            
            # Format variations to look for in text
            amount_exact = str(amount_float)
            amount_with_commas = "{:,.2f}".format(amount_float)
            
            # Extra check for special formatting if amount > 1000
            amount_patterns = [
                amount_exact, 
                amount_with_commas,
                "{:,.0f}".format(amount_float) if amount_float >= 1000 else ""
            ]
            
            # Check how many times amount appears in document
            occurrences = 0
            for pattern in amount_patterns:
                if pattern and pattern in text:
                    occurrences += 1
            
            # Valid if the amount appears at least once in the document
            return occurrences > 0
            
        except:
            return False

    def reset_bill_verification(self):
        """Reset the bill verification tab"""
        # Clear image
        self.bill_path = None
        self.bill_cv = None
        
        # Reset display
        self.bill_canvas.set_text("No document uploaded\nUse 'Upload Document' to begin verification")
        
        # Reset fields
        self.tx_id_var.set("")
        self.amount_var.set("")
        self.datetime_var.set("")
        
        # Reset document type
        self.bill_type_var.set("banking")
        
        # Clear results
        self.results_text.config(state=tk.NORMAL)
        self.results_text.delete(1.0, tk.END)
        self.results_text.config(state=tk.DISABLED)

class ZoomableCanvas(tk.Frame):
    """Canvas that supports zooming and panning"""
    def __init__(self, parent, **kwargs):
        super().__init__(parent, **kwargs)
        
        # Initialize variables
        self.zoom_level = 1.0
        self.min_zoom = 0.1
        self.max_zoom = 10.0  # Increased max zoom for more detail
        self.image = None
        self.tk_image = None
        self.original_image = None
        
        # Setup canvas
        self.canvas = tk.Canvas(self, bg="#000000", highlightthickness=0)
        self.canvas.pack(fill=tk.BOTH, expand=True)
        
        # Setup scrollbars
        self.h_scrollbar = ttk.Scrollbar(self, orient=tk.HORIZONTAL, command=self.canvas.xview)
        self.v_scrollbar = ttk.Scrollbar(self, orient=tk.VERTICAL, command=self.canvas.yview)
        
        self.canvas.configure(xscrollcommand=self.h_scrollbar.set, yscrollcommand=self.v_scrollbar.set)
        
        self.h_scrollbar.pack(fill=tk.X, side=tk.BOTTOM)
        self.v_scrollbar.pack(fill=tk.Y, side=tk.RIGHT)
        
        # Setup zoom controls
        self.zoom_frame = tk.Frame(self, bg="#f0f0f0")
        self.zoom_frame.pack(fill=tk.X, side=tk.TOP)
        
        # Improved zoom controls with additional options
        self.zoom_out_btn = tk.Button(self.zoom_frame, text="-", command=self.zoom_out, width=2)
        self.zoom_out_btn.pack(side=tk.LEFT, padx=2)
        
        self.zoom_reset_btn = tk.Button(self.zoom_frame, text="100%", command=self.reset_zoom)
        self.zoom_reset_btn.pack(side=tk.LEFT, padx=2)
        
        self.zoom_fit_btn = tk.Button(self.zoom_frame, text="Fit", command=self.fit_to_canvas)
        self.zoom_fit_btn.pack(side=tk.LEFT, padx=2)
        
        self.zoom_in_btn = tk.Button(self.zoom_frame, text="+", command=self.zoom_in, width=2)
        self.zoom_in_btn.pack(side=tk.LEFT, padx=2)
        
        self.zoom_label = tk.Label(self.zoom_frame, text="Zoom: 100%", width=10)
        self.zoom_label.pack(side=tk.RIGHT, padx=5)
        
        # Bind events
        self.canvas.bind("<MouseWheel>", self.on_mousewheel)  # Windows
        self.canvas.bind("<Button-4>", lambda e: self.on_mousewheel(e, 1))  # Linux scroll up
        self.canvas.bind("<Button-5>", lambda e: self.on_mousewheel(e, -1))  # Linux scroll down
        
        # Pan with middle-click drag or left-click drag
        self.canvas.bind("<ButtonPress-1>", self.start_pan)
        self.canvas.bind("<B1-Motion>", self.pan)
        
        # Track canvas size changes to adjust fit-to-canvas
        self.bind("<Configure>", self.on_resize)
        self.canvas_width = 0
        self.canvas_height = 0
        
    def set_image(self, image):
        """Set a new image to display"""
        if image is None:
            return
        
        self.original_image = image
        self.image = image
        
        # Default to fit-to-canvas instead of 100% for better initial view
        self.after(100, self.fit_to_canvas)  # Slight delay to ensure canvas size is updated
        
    def fit_to_canvas(self):
        """Resize the image to fit the canvas"""
        if self.original_image is None:
            return
            
        # Get canvas size (accounting for scrollbars)
        canvas_width = self.canvas.winfo_width() - 2  # Subtract border width
        canvas_height = self.canvas.winfo_height() - 2
        
        if canvas_width <= 1 or canvas_height <= 1:  # Canvas not yet drawn
            self.after(100, self.fit_to_canvas)  # Try again after a delay
            return
            
        # Calculate zoom needed to fit image
        width_ratio = canvas_width / self.original_image.width
        height_ratio = canvas_height / self.original_image.height
        
        # Use the smaller ratio to ensure the entire image fits
        self.zoom_level = min(width_ratio, height_ratio) * 0.95  # 95% of actual fit for a small margin
        
        # Update image with new zoom level
        self.image = self.original_image
        self.display_image()
        
        # Update scrollbars to center the image
        self.center_image()
        
    def center_image(self):
        """Center the image in the canvas"""
        if self.tk_image:
            # Get canvas and image dimensions
            canvas_width = self.canvas.winfo_width()
            canvas_height = self.canvas.winfo_height()
            image_width = self.tk_image.width()
            image_height = self.tk_image.height()
            
            # Calculate scroll fractions to center
            if image_width > canvas_width:
                x_center = 0.5
                self.canvas.xview_moveto(x_center - (canvas_width / 2 / image_width))
            else:
                # Image narrower than canvas, no scroll needed
                pass
                
            if image_height > canvas_height:
                y_center = 0.5
                self.canvas.yview_moveto(y_center - (canvas_height / 2 / image_height))
            else:
                # Image shorter than canvas, no scroll needed
                pass
        
    def display_image(self):
        """Update the canvas with the current image and zoom level"""
        if self.image is None:
            return
            
        # Convert numpy array to PIL Image if needed
        if isinstance(self.image, np.ndarray):
            self.image = Image.fromarray(self.image)
            
        # Calculate new dimensions based on zoom
        new_width = int(self.image.width * self.zoom_level)
        new_height = int(self.image.height * self.zoom_level)
        
        # Ensure minimum dimensions
        new_width = max(new_width, 1)
        new_height = max(new_height, 1)
        
        # Resize the image
        if self.zoom_level != 1.0:
            resized_img = self.image.resize((new_width, new_height), LANCZOS)
        else:
            resized_img = self.image
        
        # Convert to PhotoImage for canvas
        self.tk_image = ImageTk.PhotoImage(resized_img)
        
        # Update canvas
        self.canvas.delete("all")
        self.canvas.create_image(0, 0, image=self.tk_image, anchor=tk.NW, tags="img")
        self.canvas.config(scrollregion=self.canvas.bbox("all"))
        
        # Update zoom label with percentage rounded to nearest integer
        zoom_percent = int(self.zoom_level * 100)
        self.zoom_label.config(text=f"Zoom: {zoom_percent}%")
        self.zoom_reset_btn.config(text=f"{zoom_percent}%")
        
    def zoom_in(self):
        """Increase zoom level"""
        if self.zoom_level < self.max_zoom:
            # Smoother zoom increments for better control
            if self.zoom_level < 0.5:
                self.zoom_level *= 1.2
            elif self.zoom_level < 1.0:
                self.zoom_level *= 1.15
            else:
                self.zoom_level *= 1.25
                
            self.zoom_level = min(self.zoom_level, self.max_zoom)  # Ensure we don't exceed max
            self.display_image()
            
    def zoom_out(self):
        """Decrease zoom level"""
        if self.zoom_level > self.min_zoom:
            # Smoother zoom decrements for better control
            if self.zoom_level > 1.0:
                self.zoom_level /= 1.25
            elif self.zoom_level > 0.5:
                self.zoom_level /= 1.15
            else:
                self.zoom_level /= 1.1
                
            self.zoom_level = max(self.zoom_level, self.min_zoom)  # Ensure we don't go below min
            self.display_image()
            
    def reset_zoom(self):
        """Reset to original 100% zoom level"""
        self.zoom_level = 1.0
        self.display_image()
        
        # Center the image
        self.center_image()
        
    def on_mousewheel(self, event, delta=None):
        """Handle mousewheel events for zooming at cursor position"""
        if self.original_image is None:
            return
            
        # Store current scroll position and cursor position
        current_x = self.canvas.canvasx(event.x)
        current_y = self.canvas.canvasy(event.y)
        
        # Calculate cursor position relative to image
        relative_x = current_x / (self.original_image.width * self.zoom_level)
        relative_y = current_y / (self.original_image.height * self.zoom_level)
            
        # Get direction from event
        if delta is None:
            delta = event.delta
            
        # Determine zoom factor based on scroll direction
        old_zoom = self.zoom_level
        if delta > 0:
            # Zoom in - more gradual when zoomed in
            if self.zoom_level < 0.5:
                self.zoom_level *= 1.1
            elif self.zoom_level < 1.0:
                self.zoom_level *= 1.05
            else:
                self.zoom_level *= 1.1
            self.zoom_level = min(self.zoom_level, self.max_zoom)
        else:
            # Zoom out - more gradual when zoomed out
            if self.zoom_level > 1.0:
                self.zoom_level /= 1.1
            elif self.zoom_level > 0.5:
                self.zoom_level /= 1.05
            else:
                self.zoom_level /= 1.1
            self.zoom_level = max(self.zoom_level, self.min_zoom)
        
        # Apply the zoom
        self.display_image()
        
        # Calculate new scroll position to keep cursor at same relative position
        new_x = relative_x * (self.original_image.width * self.zoom_level)
        new_y = relative_y * (self.original_image.height * self.zoom_level)
        
        # Adjust scrollbars to maintain position under cursor
        self.canvas.xview_moveto((new_x - event.x) / self.tk_image.width())
        self.canvas.yview_moveto((new_y - event.y) / self.tk_image.height())
            
    def start_pan(self, event):
        """Start panning the image"""
        self.canvas.scan_mark(event.x, event.y)
        
    def pan(self, event):
        """Pan the image"""
        self.canvas.scan_dragto(event.x, event.y, gain=1)
        
    def set_text(self, text):
        """Display a text message instead of an image"""
        self.original_image = None
        self.image = None
        self.tk_image = None
        self.canvas.delete("all")
        self.canvas.create_text(
            self.canvas.winfo_width() // 2, 
            self.canvas.winfo_height() // 2, 
            text=text, 
            fill="white",
            font=("Arial", 12)
        )
        
    def on_resize(self, event):
        """Handle canvas resizing to maintain fit-to-view if needed"""
        # Only update if canvas size changed significantly
        if (abs(self.canvas_width - self.canvas.winfo_width()) > 10 or
            abs(self.canvas_height - self.canvas.winfo_height()) > 10):
            
            # Update saved dimensions
            self.canvas_width = self.canvas.winfo_width()
            self.canvas_height = self.canvas.winfo_height()
            
            # If image exists, consider refitting to canvas
            if self.original_image and hasattr(self, 'auto_fit_mode') and self.auto_fit_mode:
                self.fit_to_canvas()

def main():
    # Check if tesseract is installed and configured
    try:
        #=============================================================================
        # TESSERACT CONFIGURATION - MODIFY PATH BELOW TO MATCH YOUR INSTALLATION
        #=============================================================================
        # Set the path to your Tesseract OCR executable
        # Uncomment and modify the line below if needed:
        # pytesseract.pytesseract.tesseract_cmd = r'C:/Program Files/Tesseract-OCR/tesseract.exe'
        # 
        # For Thai language support, ensure you have installed the Thai language pack
        # and have the 'tha.traineddata' file in your Tesseract tessdata folder
        #=============================================================================
        
        # Check if Thai language data is available
        try:
            test_text = pytesseract.image_to_string(Image.new('RGB', (10, 10), color='white'), lang='tha')
        except pytesseract.TesseractError as e:
            if "Failed loading language 'tha'" in str(e):
                messagebox.showwarning(
                    "Thai Language Data Missing",
                    "Thai language pack not found. Thai text extraction will not work.\n"
                    "Please install the Thai language pack for Tesseract OCR."
                )
        
        root = tk.Tk()
        app = ImageTextComparator(root)
        root.mainloop()
    except Exception as e:
        print(f"Error starting application: {e}")
        print("Make sure Tesseract OCR is installed correctly.")

if __name__ == "__main__":
    main()