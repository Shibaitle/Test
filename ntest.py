import tkinter as tk
from tkinter import filedialog, messagebox
import cv2
import pytesseract
import Levenshtein
import os

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
        self.root.geometry("900x600")
        self.root.configure(bg="#f0f0f0")
        
        # Initialize variables
        self.image1_path = None
        self.image2_path = None
        self.image1_text = ""
        self.image2_text = ""
        
        # Language selection variables
        self.lang1_var = tk.StringVar(value="eng")  # Default to English for image 1
        self.lang2_var = tk.StringVar(value="eng")  # Default to English for image 2
        
        # Create main frame
        self.main_frame = tk.Frame(root, bg="#f0f0f0")
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
        
    def browse_image(self, image_num):
        """Open a file dialog to select an image"""
        file_path = filedialog.askopenfilename(
            title=f"Select Image {image_num}",
            filetypes=(("Image files", "*.jpg;*.jpeg;*.png;*.bmp"), ("All files", "*.*"))
        )
        
        if file_path:
            try:
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
                    self.image1_label.config(image=photo)
                    self.image1_label.image = photo  # Keep a reference
                else:
                    self.image2_path = file_path
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
            
            # Normalize text for comparison
            text1_words = self.image1_text.split()
            text2_words = self.image2_text.split()
            
            # Use difflib to find differences between words
            import difflib
            matcher = difflib.SequenceMatcher(None, text1_words, text2_words)
            
            # Display text with highlighted differences
            for tag, i1, i2, j1, j2 in matcher.get_opcodes():
                if tag == 'equal':
                    # Same text in both images - display normally
                    self.text1.insert(tk.END, ' '.join(text1_words[i1:i2]) + ' ')
                    self.text2.insert(tk.END, ' '.join(text2_words[j1:j2]) + ' ')
                elif tag == 'replace':
                    # Different text - highlight in red
                    self.text1.insert(tk.END, ' '.join(text1_words[i1:i2]) + ' ', 'diff')
                    self.text2.insert(tk.END, ' '.join(text2_words[j1:j2]) + ' ', 'diff')
                elif tag == 'delete':
                    # Text in image 1 but not in image 2
                    self.text1.insert(tk.END, ' '.join(text1_words[i1:i2]) + ' ', 'diff')
                elif tag == 'insert':
                    # Text in image 2 but not in image 1
                    self.text2.insert(tk.END, ' '.join(text2_words[j1:j2]) + ' ', 'diff')
            
            # Configure text tags for highlighting
            self.text1.tag_configure('diff', background='pink', foreground='red')
            self.text2.tag_configure('diff', background='pink', foreground='red')
            
            # Calculate and display similarity
            distance = Levenshtein.distance(' '.join(text1_words).lower(), ' '.join(text2_words).lower())
            max_len = max(len(' '.join(text1_words)), len(' '.join(text2_words)))
            similarity = ((max_len - distance) / max_len) * 100 if max_len > 0 else 100
            
            if similarity == 100:
                self.result_label.config(text="✓ The text in both images is the same.", fg="green")
            else:
                self.result_label.config(
                    text=f"✗ The text in the images is different. (Similarity: {similarity:.1f}%)", 
                    fg="red"
                )
                
        except Exception as e:
            messagebox.showerror("Error", f"Comparison failed: {e}")
    
    def reset(self):
        """Reset the application to its initial state"""
        # Clear images
        self.image1_path = None
        self.image2_path = None
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
        
        # Clear stored text
        self.image1_text = ""
        self.image2_text = ""

def main():
    # Check if tesseract is installed and configured
    try:
        #=============================================================================
        # TESSERACT CONFIGURATION - MODIFY PATH BELOW TO MATCH YOUR INSTALLATION
        #=============================================================================
        # Set the path to your Tesseract OCR executable
        # **
        ## pytesseract.pytesseract.tesseract_cmd = r'C:/Program Files/Tesseract-OCR/tesseract.exe'
        # **
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