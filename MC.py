def get_manual_content():
    """Returns the complete user manual content for the Image Text Comparison Tool"""
    
    manual_text = """
IMAGE TEXT COMPARISON TOOL - USER MANUAL
========================================

VERSION: 2.0
LAST UPDATED: 2025

TABLE OF CONTENTS
================
1. Introduction
2. Getting Started
3. Text Comparison Tab
4. Visual Comparison Tab
5. Advanced Detection Tab
6. Language Support
7. Troubleshooting
8. Tips for Best Results

1. INTRODUCTION
===============
The Image Text Comparison Tool is a comprehensive application designed to compare text content between two images and detect potential document forgeries or manipulations. The tool uses advanced OCR (Optical Character Recognition) technology and image analysis algorithms to provide accurate text extraction and comparison results.

Key Features:
- Text extraction from images using Tesseract OCR
- Support for English and Thai languages
- Side-by-side text comparison with highlighting
- Visual difference detection
- Advanced counterfeit detection algorithms
- Zoomable image viewer with pan controls

2. GETTING STARTED
==================

System Requirements:
- Python 3.7 or higher
- Tesseract OCR installed on your system
- Required Python packages: opencv-python, pytesseract, Pillow, tkinter, numpy, scikit-learn, Levenshtein

Installation:
1. Install Tesseract OCR from: https://github.com/tesseract-ocr/tesseract
2. For Thai language support, ensure the Thai language pack is installed
3. Install required Python packages using pip
4. Run the application by executing: python ntestnew1.py

First Launch:
- The application will check for Tesseract installation
- If Thai language support is missing, you'll see a warning
- The main window will open with three tabs available

3. TEXT COMPARISON TAB
======================

This is the main tab for comparing text content between two images.

3.1 Loading Images:
- Click "Browse" button under Image 1 or Image 2
- Select image files (supported formats: JPG, JPEG, PNG, BMP)
- Images will be displayed with automatic resizing for optimal viewing
- Both images must be loaded before comparison

3.2 Language Selection:
- Each image has its own language dropdown
- Available options:
  * eng: English only
  * tha: Thai only
  * tha+eng: Thai and English combined
- Select the appropriate language before text extraction

3.3 Text Extraction:
- Click "Extract Text" for each image to perform OCR
- Text will appear in the text areas below the images
- The tool automatically preprocesses images for better OCR accuracy
- Line numbers are added for easy reference

3.4 Text Comparison:
- Click "Compare Text" to analyze both extracted texts
- Differences are highlighted in pink/red
- Similarity percentage is calculated and displayed
- Results show:
  * Overall similarity score
  * Line-by-line differences
  * Word-level highlighting of changes

3.5 Action Buttons:
- Compare Text: Performs detailed text comparison
- Show Visual Diff: Creates visual difference image (switches to Visual tab)
- User Manual: Opens this manual window
- Reset: Clears all data and returns to initial state

4. VISUAL COMPARISON TAB
========================

This tab provides visual analysis of differences between images.

4.1 Visual Difference Display:
- Shows a side-by-side comparison: Image 1 | Differences | Image 2
- Red areas indicate significant differences between images
- Uses advanced image processing to detect visual changes
- Zoomable canvas with mouse wheel support

4.2 Zoom Controls:
- Mouse wheel: Zoom in/out at cursor position
- "+" button: Zoom in
- "-" button: Zoom out
- "100%" button: Reset to original size
- "Fit" button: Fit image to window size

4.3 Pan Controls:
- Left-click and drag to pan around the image
- Scroll bars appear when image is larger than view area

4.4 Best Practices:
- Use images of similar size for better comparison
- Ensure images have similar content alignment
- Higher resolution images provide better difference detection

5. ADVANCED DETECTION TAB
==========================

This tab provides sophisticated counterfeit detection capabilities.

5.1 Detection Controls:
- Sensitivity Slider (1-100%): Adjusts detection sensitivity
  * Higher values detect more subtle differences
  * Lower values focus on major differences only
- Detection Method:
  * Contour: Detects shape and edge differences
  * Feature Match: Compares distinctive image features

5.2 Detection Process:
1. Load both images in the Text Comparison tab
2. Switch to Advanced Detection tab
3. Adjust sensitivity as needed
4. Select detection method
5. Click "Detect Differences"

5.3 Results Display:
- Red bounding boxes highlight suspicious areas
- Detection details show:
  * Number of suspicious areas found
  * Confidence score (0-100%)
  * Risk assessment (Low/Medium/High)
- Side-by-side view with annotations

5.4 Risk Assessment:
- HIGH RISK (75%+): Significant manipulations detected
- MEDIUM RISK (30-75%): Some suspicious differences
- LOW RISK (<30%): Minor differences only

6. LANGUAGE SUPPORT
===================

6.1 English (eng):
- Optimized for Western text
- Best for documents with Latin characters
- Includes number and symbol recognition

6.2 Thai (tha):
- Specialized for Thai script
- Handles complex Thai character combinations
- Optimized for Thai document formats

6.3 Thai + English (tha+eng):
- Dual language detection
- Best for mixed-language documents
- Slightly slower processing time
- Recommended for most Thai business documents

6.4 OCR Optimization:
- Images are automatically preprocessed for better accuracy
- Noise reduction and contrast enhancement applied
- Automatic deskewing for tilted text
- Resolution upscaling for small images

7. TROUBLESHOOTING
==================

7.1 Common Issues:

"Tesseract not found" error:
- Ensure Tesseract OCR is properly installed
- Check the installation path in the code
- Verify Tesseract is in your system PATH

Poor text extraction quality:
- Use higher resolution images (minimum 300 DPI recommended)
- Ensure good contrast between text and background
- Try different language settings
- Check for image skew or rotation

"Thai language pack missing" warning:
- Download and install Thai language data for Tesseract
- Place 'tha.traineddata' in Tesseract's tessdata folder
- Restart the application after installation

Memory issues with large images:
- Resize images to reasonable dimensions (max 3000x3000 pixels)
- Close other applications to free memory
- Use smaller images for initial testing

7.2 Performance Tips:
- Process images one at a time for large files
- Use SSD storage for better file access speed
- Ensure adequate RAM (4GB+ recommended)
- Close unnecessary background applications

8. TIPS FOR BEST RESULTS
========================

8.1 Image Quality:
- Use clear, high-contrast images
- Avoid blurry or low-resolution photos
- Ensure even lighting across the document
- Minimize shadows and reflections

8.2 Document Preparation:
- Scan documents at 300 DPI or higher
- Keep documents flat during scanning
- Use white or light backgrounds
- Avoid handwritten annotations over printed text

8.3 Comparison Strategy:
1. Start with Text Comparison for content analysis
2. Use Visual Comparison for layout differences
3. Apply Advanced Detection for forgery analysis
4. Cross-reference results from all three methods

8.4 Language Selection Guidelines:
- Use 'eng' for pure English documents
- Use 'tha' for pure Thai documents
- Use 'tha+eng' for mixed documents or when uncertain
- Test different language settings if results are poor

8.5 Detection Sensitivity:
- Start with 50% sensitivity for general use
- Increase to 70-80% for detailed forensic analysis
- Decrease to 20-30% for major difference detection only
- Adjust based on image quality and content type

SUPPORT AND UPDATES
===================
For technical support or to report issues, please refer to the project documentation or contact the development team.

This tool is designed for educational and professional document verification purposes. Always combine automated analysis with human expertise for critical decisions.

Copyright © 2025 - Image Text Comparison Tool
All rights reserved.
"""
    
    return manual_text