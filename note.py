# Document Verification Tool: Detection Feature Guide
# 
# About the Detection Feature
# The Advanced Detection tab in our Image Text Comparison Tool provides powerful capabilities to identify potential manipulations between documents or images. This guide explains how to use the counterfeit detection features effectively.
#
# <img alt="Detection Feature" src="https://example.com/detection-feature.png">
# 
# How It Works
# The detect_counterfeit function analyzes two images to identify areas that differ significantly, which could indicate tampering or manipulation. The tool uses computer vision techniques to highlight suspicious regions with red bounding boxes.
#
# Detection Methods
# 
# 1. Contour Detection
# - Best for: Detecting structural changes, added/removed elements, or text modifications
# - How it works: Identifies outline differences between images by comparing pixel values
# - When to use: When comparing documents where text or elements might have been added, removed, or modified
#
# 2. Feature Matching
# - Best for: Detecting subtle manipulations, image replacements, or sophisticated edits
# - How it works: Identifies key visual elements (features) that don't match properly between images
# - When to use: When comparing images where parts might have been replaced or manipulated
# 
# Using the Detection Feature
# 1. Upload Images: Select two images to compare (original and potentially modified document)
# 2. Navigate: Go to the "Advanced Detection" tab
# 3. Adjust Settings:
#    - Sensitivity: Control how aggressively differences are detected (higher = more sensitive)
#    - Method: Choose between Contour or Feature Match based on your needs
# 4. Click "Detect Differences": The tool will analyze and highlight suspicious areas
# 5. Review Results: Examine the highlighted areas and review the risk assessment
#
# Understanding Results
# 
# Visual Output
# - Red boxes highlight suspicious areas on both images
# - Labels mark each potentially manipulated region
# 
# Detection Details
# - Confidence Score: Indicates the likelihood of manipulation (0-100%)
# - Risk Assessment:
#   - HIGH RISK: Significant manipulations detected (score > 75%)
#   - MEDIUM RISK: Some suspicious differences detected (score 30-75%)
#   - LOW RISK: Minor differences detected (score < 30%)
#
# Tips for Accurate Detection
# 1. Adjust Sensitivity: Start with 50% and increase if needed
# 2. Try Both Methods: Feature matching may catch things contour detection misses and vice versa
# 3. Similar Sizes: For best results, use images of similar dimensions
# 4. Clean Documents: Better results with clearly visible, high-contrast documents
# 5. Verify Manually: Always manually verify suspicious areas highlighted by the tool
#
# Limitations
# - May produce false positives with poorly aligned documents
# - Highly compressed or low-quality images may reduce detection accuracy
# - Complex backgrounds can interfere with accurate detection
#
# Practical Applications
# - Verifying official documents against known originals
# - Detecting unauthorized modifications in contracts or agreements
# - Identifying manipulated photos or images
# - Quality control of printed materials
#
# For technical details or to learn more about the algorithms used, please contact the development team.