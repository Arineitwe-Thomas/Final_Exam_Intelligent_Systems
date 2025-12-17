"""
Script to extract images from Jupyter notebook and insert them into Word document
"""

import json
import base64
import os
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.style import WD_STYLE_TYPE
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
import re

def extract_images_from_notebook(notebook_path, output_dir):
    """Extract all images from notebook outputs and save them"""
    with open(notebook_path, 'r', encoding='utf-8') as f:
        notebook = json.load(f)
    
    images = []
    image_counter = 0
    
    for cell_idx, cell in enumerate(notebook.get('cells', [])):
        if cell.get('cell_type') == 'code' and 'outputs' in cell:
            for output_idx, output in enumerate(cell['outputs']):
                if output.get('output_type') == 'display_data' and 'image/png' in output.get('data', {}):
                    image_data = output['data']['image/png']
                    image_counter += 1
                    image_filename = f"image_{image_counter:02d}.png"
                    image_path = os.path.join(output_dir, image_filename)
                    
                    # Decode base64 and save
                    with open(image_path, 'wb') as img_file:
                        img_file.write(base64.b64decode(image_data))
                    
                    # Store metadata
                    images.append({
                        'path': image_path,
                        'filename': image_filename,
                        'cell_idx': cell_idx,
                        'output_idx': output_idx
                    })
                    print(f"Extracted image {image_counter}: {image_filename}")
    
    return images

def get_cell_source_text(cell):
    """Get source text from a notebook cell"""
    source = cell.get('source', [])
    if isinstance(source, list):
        return ''.join(source)
    return str(source)

def identify_image_context(notebook, image_info):
    """Identify what the image represents based on cell content"""
    cell = notebook['cells'][image_info['cell_idx']]
    source = get_cell_source_text(cell).lower()
    
    # Identify image type based on cell content
    if 'silhouette' in source and 'k-means' in source:
        return 'kmeans_silhouette'
    elif 'knn distance' in source or 'eps' in source:
        return 'dbscan_knn'
    elif 'pca' in source and 'cluster' in source and 'kmeans' in source:
        return 'kmeans_pca'
    elif 'pca' in source and 'cluster' in source and 'dbscan' in source:
        return 'dbscan_pca'
    elif 'autoencoder' in source and 'summary' in source:
        return 'autoencoder_summary'
    elif 'autoencoder' in source and 'loss' in source:
        return 'autoencoder_loss'
    elif 'latent' in source and 'embedding' in source:
        return 'autoencoder_latent'
    elif 'pca' in source and 'autoencoder' in source and 'comparison' in source:
        return 'pca_vs_ae'
    elif 'fpgrowth' in source or 'frequent' in source:
        return 'fpgrowth'
    elif 'association' in source and 'rule' in source:
        return 'association_rules'
    else:
        return 'unknown'

def create_report_with_images():
    """Create Word document with images inserted"""
    
    notebook_path = r"c:\INT_SYSTEMS\Final exam\FINAL_EXAM_INTELLIGENT_SYSTEM\Backup\Question 2\Customer Purchasing Behavior.ipynb"
    output_dir = r"c:\INT_SYSTEMS\Final exam\FINAL_EXAM_INTELLIGENT_SYSTEM\Backup\Question 2\images"
    
    # Create images directory
    os.makedirs(output_dir, exist_ok=True)
    
    # Extract images from notebook
    print("Extracting images from notebook...")
    images = extract_images_from_notebook(notebook_path, output_dir)
    
    # Map images to contexts
    with open(notebook_path, 'r', encoding='utf-8') as f:
        notebook = json.load(f)
    
    image_map = {}
    for img in images:
        context = identify_image_context(notebook, img)
        if context not in image_map:
            image_map[context] = []
        image_map[context].append(img['path'])
    
    print(f"\nExtracted {len(images)} images")
    print(f"Image contexts: {list(image_map.keys())}")
    
    # Now create the Word document with images
    # (I'll import the create_report function and modify it)
    from generate_report import create_report_structure
    
    # Create document
    doc = Document()
    
    # Set document margins
    sections = doc.sections
    for section in sections:
        section.top_margin = Inches(1)
        section.bottom_margin = Inches(1)
        section.left_margin = Inches(1)
        section.right_margin = Inches(1)
    
    # Import the report creation logic
    # For now, let me create a modified version that inserts images
    return create_modified_report(doc, image_map, output_dir)

def create_modified_report(doc, image_map, image_dir):
    """Create report with images inserted at appropriate locations"""
    
    # This will be a comprehensive function that creates the full report
    # For brevity, I'll create a script that modifies the existing generate_report.py
    # to insert images instead of placeholders
    
    print("\nCreating Word document with images...")
    print("Note: This will modify the existing report generation to include images")
    
    # We'll need to modify generate_report.py to accept image paths and insert them
    return True

if __name__ == "__main__":
    create_report_with_images()

