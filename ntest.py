import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import cv2
import pytesseract
import Levenshtein
import os
import numpy as np

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
        
        # Setup text comparison tab
        self.setup_text_tab()
        
        # Setup visual comparison tab
        self.setup_visual_tab()
        
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
        
        # Buttons for image 1
        self.btn_frame1 = tk.Frame(self.left_frame, bg="#f0f0f0")
        self.btn_frame1.pack(fill=tk.X, pady=5)
        self.browse_btn1 = tk.Button(self.btn_frame1, text="Browse", command=lambda: self.browse_image(1))
        self.browse_btn1.pack(side=tk.LEFT, padx=5)
        
        # Language selection for image 1
        self.lang_frame1 = tk.Frame(self.btn_frame1, bg="#f0f0f0")
        self.lang_frame1.pack(side=tk.RIGHT, padx=5)
        self.eng_radio1 = tk.Radiobutton(self.lang_frame1, text="English", variable=self.lang1_var, value="eng", bg="#f0f0f0")
        self.eng_radio1.pack(side=tk.LEFT)
        self.thai_radio1 = tk.Radiobutton(self.lang_frame1, text="Thai", variable=self.lang1_var, value="tha", bg="#f0f0f0")
        self.thai_radio1.pack(side=tk.LEFT)
        
        # Buttons for image 2
        self.btn_frame2 = tk.Frame(self.right_frame, bg="#f0f0f0")
        self.btn_frame2.pack(fill=tk.X, pady=5)
        self.browse_btn2 = tk.Button(self.btn_frame2, text="Browse", command=lambda: self.browse_image(2))
        self.browse_btn2.pack(side=tk.LEFT, padx=5)
        
        # Language selection for image 2
        self.lang_frame2 = tk.Frame(self.btn_frame2, bg="#f0f0f0")
        self.lang_frame2.pack(side=tk.RIGHT, padx=5)
        self.eng_radio2 = tk.Radiobutton(self.lang_frame2, text="English", variable=self.lang2_var, value="eng", bg="#f0f0f0")
        self.eng_radio2.pack(side=tk.LEFT)
        self.thai_radio2 = tk.Radiobutton(self.lang_frame2, text="Thai", variable=self.lang2_var, value="tha", bg="#f0f0f0")
        self.thai_radio2.pack(side=tk.LEFT)
        
        # Action buttons frame
        self.action_frame = tk.Frame(self.main_frame, bg="#f0f0f0")
        self.action_frame.pack(fill=tk.X, pady=10)
        
        # Compare button
        self.compare_btn = tk.Button(self.action_frame, text="Compare Text", command=self.compare_text, bg="#4CAF50", fg="white", padx=10, pady=5)
        self.compare_btn.pack(side=tk.LEFT, padx=5)
        
        # Visual diff button
        self.visual_diff_btn = tk.Button(self.action_frame, text="Show Visual Diff", command=self.show_visual_diff, bg="#2196F3", fg="white", padx=10, pady=5)
        self.visual_diff_btn.pack(side=tk.LEFT, padx=5)
        
        # Reset button
        self.reset_btn = tk.Button(self.action_frame, text="Reset", command=self.reset, bg="#f44336", fg="white", padx=10, pady=5)
        self.reset_btn.pack(side=tk.LEFT, padx=5)
        
        # Result frame
        self.result_frame = tk.LabelFrame(self.main_frame, text="Comparison Result", bg="#f0f0f0", padx=10, pady=10)
        self.result_frame.pack(fill=tk.X, pady=10)
        
        self.result_label = tk.Label(self.result_frame, text="No comparison performed yet", bg="#f0f0f0", font=("Arial", 12))
        self.result_label.pack(pady=10)
        
        # Text display frame
        self.text_frame = tk.Frame(self.main_frame, bg="#f0f0f0")
        self.text_frame.pack(fill=tk.BOTH, expand=True, pady=5)
        
        # Text area for Image 1
        self.text_frame1 = tk.LabelFrame(self.text_frame, text="Text from Image 1", bg="#f0f0f0")
        self.text_frame1.grid(row=0, column=0, padx=5, pady=5, sticky="nsew")
        self.text1 = tk.Text(self.text_frame1, height=5, width=40)
        self.text1.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        # Text area for Image 2
        self.text_frame2 = tk.LabelFrame(self.text_frame, text="Text from Image 2", bg="#f0f0f0")
        self.text_frame2.grid(row=0, column=1, padx=5, pady=5, sticky="nsew")
        self.text2 = tk.Text(self.text_frame2, height=5, width=40)
        self.text2.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        # Configure grid weights for text frames
        self.text_frame.grid_columnconfigure(0, weight=1)
        self.text_frame.grid_columnconfigure(1, weight=1)
        self.text_frame.grid_rowconfigure(0, weight=1)
    
    def setup_visual_tab(self):
        """Setup the visual comparison tab"""
        # Create main frame for visual comparison
        self.visual_main_frame = tk.Frame(self.visual_tab, bg="#f0f0f0")
        self.visual_main_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=20)
        
        # Create frames for visual comparison
        self.visual_frames_container = tk.Frame(self.visual_main_frame, bg="#f0f0f0")
        self.visual_frames_container.pack(fill=tk.BOTH, expand=True)
        
        # Create frame for visual difference
        self.visual_diff_frame = tk.LabelFrame(self.visual_frames_container, text="Visual Difference Comparison", bg="#f0f0f0", padx=10, pady=10)
        self.visual_diff_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # Create image label for visual difference
        self.visual_diff_label = tk.Label(self.visual_diff_frame, text="No visual comparison available\nClick 'Show Visual Diff' button to generate", 
                                        bg="#e0e0e0", width=80, height=30)
        self.visual_diff_label.pack(fill=tk.BOTH, expand=True)
        
        # Instructions
        self.visual_instructions = tk.Label(self.visual_main_frame, 
                                         text="Red areas indicate differences between the images.\nFor best results, use similar-sized images with similar content.",
                                         bg="#f0f0f0", font=("Arial", 10))
        self.visual_instructions.pack(pady=10)

    def browse_image(self, image_num):
        """Open a file dialog to select an image"""
        file_path = filedialog.askopenfilename(
            title=f"Select Image {image_num}",
            filetypes=(("Image files", "*.jpg;*.jpeg;*.png;*.bmp"), ("All files", "*.*"))
        )
        
        if file_path:
            try:
                # Store the original image for processing
                cv_img = cv2.imread(file_path)
                
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
        """Extract text from an image using pytesseract OCR"""
        if not image_path:
            return ""
        
        try:
            # Read image with OpenCV
            img = cv2.imread(image_path)
            
            # Convert to grayscale
            gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
            
            # Apply different processing based on language
            if lang == "tha":
                # For Thai: use adaptive thresholding which works better for complex scripts
                binary = cv2.adaptiveThreshold(
                    gray, 255, cv2.ADAPTIVE_THRESH_GAUSSIAN_C,
                    cv2.THRESH_BINARY, 11, 2
                )
            else:
                # For English: use regular thresholding
                _, binary = cv2.threshold(gray, 150, 255, cv2.THRESH_BINARY)
            
            # Use pytesseract to extract text with specified language
            text = pytesseract.image_to_string(binary, lang=lang)
            
            # Return text but preserve line breaks (only strip trailing/leading whitespace)
            return text.strip()
        
        except Exception as e:
            messagebox.showerror("Error", f"Failed to extract text: {e}")
            return ""
    
    def compare_text(self):
        """Extract text from both images and compare them"""
        if not self.image1_path or not self.image2_path:
            messagebox.showwarning("Warning", "Please select both images first!")
            return
        
        try:
            # Extract text using selected language for each image
            self.image1_text = self.extract_text(self.image1_path, self.lang1_var.get())
            self.image2_text = self.extract_text(self.image2_path, self.lang2_var.get())
            
            # Clear previous text
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
            
            # Display each line separately, with line numbers
            for i, line in enumerate(text1_lines, 1):
                self.text1.insert(tk.END, f"{i}: {line}\n")
            
            for i, line in enumerate(text2_lines, 1):
                self.text2.insert(tk.END, f"{i}: {line}\n")
            
            # Calculate and display similarity by joining all text
            text1_joined = ' '.join(text1_lines).lower()
            text2_joined = ' '.join(text2_lines).lower()
            
            # Compare using Levenshtein distance
            distance = Levenshtein.distance(text1_joined, text2_joined)
            max_len = max(len(text1_joined), len(text2_joined))
            similarity = ((max_len - distance) / max_len) * 100 if max_len > 0 else 100
            
            # Compare line by line
            num_lines = min(len(text1_lines), len(text2_lines))
            diff_lines = []
            for i in range(num_lines):
                if text1_lines[i].strip() != text2_lines[i].strip():
                    diff_lines.append(i+1)  # +1 because we use 1-based indexing for display
            
            # Always display similarity score regardless of result
            if len(text1_lines) != len(text2_lines):
                self.result_label.config(
                    text=f"✗ Different number of lines: Image 1 ({len(text1_lines)}) vs Image 2 ({len(text2_lines)}) (Similarity: {similarity:.1f}%)", 
                    fg="red"
                )
            elif diff_lines:
                self.result_label.config(
                    text=f"✗ Differences found in lines: {', '.join(map(str, diff_lines))} (Similarity: {similarity:.1f}%)", 
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
            
            # Highlight different lines
            for line_num in diff_lines:
                # Calculate the position in the Text widget for start and end of this line
                start_pos1 = f"{line_num}.0"
                end_pos1 = f"{line_num}.end"
                start_pos2 = f"{line_num}.0"
                end_pos2 = f"{line_num}.end"
                
                # Apply the highlighting tag to this line
                self.text1.tag_add('diff', start_pos1, end_pos1)
                self.text2.tag_add('diff', start_pos2, end_pos2)
            
            # Configure text tags for highlighting
            self.text1.tag_configure('diff', background='pink', foreground='red')
            self.text2.tag_configure('diff', background='pink', foreground='red')
                
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
            
            # Resize for display in the UI
            display_height = 500
            display_width = int(comparison_width * (display_height / max_height))
            comparison_display = cv2.resize(comparison_img, (display_width, display_height))
            
            # Convert to PIL format for Tkinter
            comparison_pil = Image.fromarray(comparison_display)
            comparison_photo = ImageTk.PhotoImage(comparison_pil)
            
            # Update the visual diff label
            self.visual_diff_label.config(image=comparison_photo, text='')
            self.visual_diff_label.image = comparison_photo  # Keep a reference
            
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
        
        # Reset visual diff
        self.visual_diff_label.config(image="", text="No visual comparison available\nClick 'Show Visual Diff' button to generate")
        
        # Clear stored text
        self.image1_text = ""
        self.image2_text = ""
        
        # Switch to text tab
        self.notebook.select(self.text_tab)

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