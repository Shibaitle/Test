import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import cv2
import pytesseract
import Levenshtein
import os
import numpy as np
from sklearn.cluster import DBSCAN

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
        
        # Setup text comparison tab
        self.setup_text_tab()
        
        # Setup visual comparison tab
        self.setup_visual_tab()
        
        # Setup counterfeit detection tab (new)
        self.setup_detection_tab()
        
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

class ZoomableCanvas(tk.Frame):
    """Canvas that supports zooming and panning"""
    def __init__(self, parent, **kwargs):
        super().__init__(parent, **kwargs)
        
        # Initialize variables
        self.zoom_level = 1.0
        self.min_zoom = 0.1
        self.max_zoom = 5.0
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
        
        self.zoom_in_btn = tk.Button(self.zoom_frame, text="+", command=self.zoom_in)
        self.zoom_in_btn.pack(side=tk.LEFT, padx=2)
        
        self.zoom_out_btn = tk.Button(self.zoom_frame, text="-", command=self.zoom_out)
        self.zoom_out_btn.pack(side=tk.LEFT, padx=2)
        
        self.zoom_reset_btn = tk.Button(self.zoom_frame, text="Reset Zoom", command=self.reset_zoom)
        self.zoom_reset_btn.pack(side=tk.LEFT, padx=2)
        
        self.zoom_label = tk.Label(self.zoom_frame, text="Zoom: 100%", width=10)
        self.zoom_label.pack(side=tk.LEFT, padx=5)
        
        # Bind events
        self.canvas.bind("<MouseWheel>", self.on_mousewheel)  # Windows
        self.canvas.bind("<Button-4>", lambda e: self.on_mousewheel(e, 1))  # Linux scroll up
        self.canvas.bind("<Button-5>", lambda e: self.on_mousewheel(e, -1))  # Linux scroll down
        
        # Pan with middle-click drag or left-click drag
        self.canvas.bind("<ButtonPress-1>", self.start_pan)
        self.canvas.bind("<B1-Motion>", self.pan)
        
    def set_image(self, image):
        """Set a new image to display"""
        if image is None:
            return
        
        self.original_image = image
        self.image = image
        self.display_image()
        self.reset_zoom()
        
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
        
        # Update zoom label
        self.zoom_label.config(text=f"Zoom: {int(self.zoom_level * 100)}%")
        
    def zoom_in(self):
        """Increase zoom level"""
        if self.zoom_level < self.max_zoom:
            self.zoom_level *= 1.25
            self.display_image()
            
    def zoom_out(self):
        """Decrease zoom level"""
        if self.zoom_level > self.min_zoom:
            self.zoom_level /= 1.25
            self.display_image()
            
    def reset_zoom(self):
        """Reset to original zoom level"""
        self.zoom_level = 1.0
        self.display_image()
        
        # Reset scroll position to top-left
        self.canvas.xview_moveto(0)
        self.canvas.yview_moveto(0)
        
    def on_mousewheel(self, event, delta=None):
        """Handle mousewheel events for zooming"""
        if self.original_image is None:
            return
            
        # Get direction from event
        if delta is None:
            delta = event.delta
            
        # Determine zoom factor based on scroll direction
        if delta > 0:
            self.zoom_in()  # Zoom in
        else:
            self.zoom_out()  # Zoom out
            
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