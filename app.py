import streamlit as st
from PIL import Image
import io
import base64
from pathlib import Path
import pandas as pd
import json
import zipfile
from docx import Document
from docx.shared import Inches
import PyPDF2
from pdf2image import convert_from_bytes
import img2pdf
from pptx import Presentation
from pptx.util import Inches as PptxInches
import openpyxl
from reportlab.lib.pagesizes import letter, landscape
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Image as RLImage, Table, TableStyle, PageBreak
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import inch
from reportlab.lib import colors
from reportlab.lib.enums import TA_CENTER, TA_LEFT, TA_RIGHT, TA_JUSTIFY
from reportlab.lib.utils import ImageReader

# Set page configuration
st.set_page_config(
    page_title="Universal File Converter Pro",
    page_icon="🔄",
    layout="wide"
)

# Custom CSS
st.markdown("""
<style>
    .conversion-card {
        padding: 20px;
        border-radius: 10px;
        background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
        color: white;
        text-align: center;
        margin: 10px;
    }
    .stButton>button {
        width: 100%;
        background-color: #4CAF50;
        color: white;
        padding: 10px;
        font-size: 16px;
        border-radius: 5px;
    }
</style>
""", unsafe_allow_html=True)

# Title
st.title("🔄 Universal File Converter Pro")
st.markdown("### Convert any file format with ease!")

# Conversion categories
conversion_categories = {
    "📄 To PDF": [
        "Word to PDF",
        "Excel to PDF", 
        "PowerPoint to PDF",
        "JPG to PDF",
        "PNG to PDF",
        "Text to PDF"
    ],
    "📝 From PDF": [
        "PDF to Word",
        "PDF to Excel",
        "PDF to PowerPoint",
        "PDF to JPG",
        "PDF to PNG",
        "Extract PDF Images",
        "PDF to Text"
    ],
    "🛠️ PDF Tools": [
        "Merge PDF",
        "Split PDF",
        "Compress PDF",
        "Rotate PDF",
        "Remove PDF Pages",
        "Extract PDF Pages"
    ],
    "🖼️ Image Conversion": [
        "JPG to PNG",
        "PNG to JPG",
        "Image to WebP",
        "WebP to JPG",
        "WebP to PNG",
        "Image to BMP",
        "BMP to JPG",
        "Resize Image",
        "Rotate Image"
    ],
    "📊 Office Files": [
        "Word to Excel",
        "Excel to Word",
        "CSV to Excel",
        "Excel to CSV",
        "JSON to Excel",
        "Excel to JSON"
    ]
}

# Sidebar
st.sidebar.header("🎯 Select Conversion Type")
category = st.sidebar.selectbox("Category", list(conversion_categories.keys()))
conversion_type = st.sidebar.selectbox("Conversion", conversion_categories[category])

# Helper Functions

def word_to_pdf(uploaded_file):
    """Convert Word to PDF with formatting preserved"""
    try:
        from docx import Document
        
        doc = Document(uploaded_file)
        
        output = io.BytesIO()
        pdf_doc = SimpleDocTemplate(output, pagesize=letter,
                                   leftMargin=inch, rightMargin=inch,
                                   topMargin=inch, bottomMargin=inch)
        
        styles = getSampleStyleSheet()
        story = []
        
        # Create custom styles
        for i in range(1, 10):
            style_name = f'Heading{i}'
            if style_name not in styles:
                styles.add(ParagraphStyle(
                    name=style_name,
                    parent=styles['Heading1'],
                    fontSize=20 - (i * 2),
                    spaceAfter=12,
                    spaceBefore=12
                ))
        
        # Process paragraphs
        for para in doc.paragraphs:
            if not para.text.strip():
                story.append(Spacer(1, 6))
                continue
            
            style = styles['Normal']
            if para.style.name.startswith('Heading'):
                try:
                    level = para.style.name.replace('Heading', '').strip()
                    if level.isdigit():
                        style = styles[f'Heading{level}']
                    else:
                        style = styles['Heading1']
                except:
                    style = styles['Heading1']
            elif 'Title' in para.style.name:
                style = styles['Title']
            
            text = para.text
            p = Paragraph(text, style)
            story.append(p)
            story.append(Spacer(1, 6))
        
        # Process tables
        for table in doc.tables:
            data = []
            for row in table.rows:
                row_data = [cell.text for cell in row.cells]
                data.append(row_data)
            
            if data:
                t = Table(data)
                t.setStyle(TableStyle([
                    ('BACKGROUND', (0, 0), (-1, 0), colors.grey),
                    ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
                    ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
                    ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
                    ('GRID', (0, 0), (-1, -1), 1, colors.black)
                ]))
                story.append(t)
                story.append(Spacer(1, 12))
        
        # Safety check: limit total story elements to prevent infinite pages
        if len(story) > 500:  # Limit total elements
            story = story[:500]
            story.append(Paragraph("... (Content truncated for safety)", styles['Normal']))
        
        try:
            pdf_doc.build(story)
        except Exception as layout_error:
            # Fallback: create simple text-only PDF
            st.warning("Complex layout failed, creating simplified text-only PDF...")
            
            # Extract all text content from Word document
            simple_story = []
            simple_story.append(Paragraph("Document Content", styles['Heading1']))
            simple_story.append(Spacer(1, 12))
            
            for para in doc.paragraphs:
                if para.text.strip():
                    # Simple text extraction only
                    text = para.text.strip()[:500]  # Limit text length
                    simple_story.append(Paragraph(text, styles['Normal']))
                    simple_story.append(Spacer(1, 6))
            
            # Build simple version
            pdf_doc.build(simple_story[:30])  # Very limited content
        
        output.seek(0)
        return output
    except Exception as e:
        st.error(f"Error: {str(e)}")
        return None

def excel_to_pdf(uploaded_file):
    """Convert Excel to PDF"""
    try:
        wb = openpyxl.load_workbook(uploaded_file)
        
        output = io.BytesIO()
        pdf_doc = SimpleDocTemplate(output, pagesize=letter)
        styles = getSampleStyleSheet()
        story = []
        
        for sheet_name in wb.sheetnames:
            ws = wb[sheet_name]
            
            story.append(Paragraph(f"<b>Sheet: {sheet_name}</b>", styles['Heading1']))
            story.append(Spacer(1, 12))
            
            data = []
            for row in ws.iter_rows():
                row_data = [str(cell.value if cell.value is not None else "") for cell in row]
                data.append(row_data)
            
            if data:
                t = Table(data, repeatRows=1)
                t.setStyle(TableStyle([
                    ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#4472C4')),
                    ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
                    ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
                    ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
                    ('GRID', (0, 0), (-1, -1), 0.5, colors.grey),
                ]))
                story.append(t)
                story.append(Spacer(1, 24))
        
        pdf_doc.build(story)
        output.seek(0)
        return output
    except Exception as e:
        st.error(f"Error: {str(e)}")
        return None

def ppt_to_pdf(uploaded_file):
    """Convert PowerPoint to PDF with enhanced formatting and image extraction"""
    try:
        # Validate file size
        file_content = uploaded_file.read()
        if len(file_content) > 50 * 1024 * 1024:  # 50MB limit
            raise ValueError("File size too large. Please upload a file smaller than 50MB.")
        
        # Reset file pointer and load presentation
        uploaded_file.seek(0)
        prs = Presentation(uploaded_file)
        
        # Validate slide count
        if len(prs.slides) == 0:
            raise ValueError("The PowerPoint file appears to be empty. Please check the file.")
        
        # Get presentation dimensions for better page sizing
        try:
            # Try to get slide dimensions from presentation
            slide_width = prs.slide_width
            slide_height = prs.slide_height
            # Convert from EMU to inches (1 inch = 914400 EMU)
            width_inches = slide_width / 914400
            height_inches = slide_height / 914400
            
            # Create custom page size based on slide dimensions with some margins
            custom_pagesize = (width_inches * inch, height_inches * inch)
        except:
            # Fallback to letter size if dimensions can't be determined
            custom_pagesize = letter
        
        output = io.BytesIO()
        pdf_doc = SimpleDocTemplate(
            output,
            pagesize=custom_pagesize,
            leftMargin=0.5*inch,
            rightMargin=0.5*inch,
            topMargin=0.5*inch,
            bottomMargin=0.5*inch
        )
        styles = getSampleStyleSheet()
        story = []
        
        # Create enhanced custom styles with better formatting
        slide_title_style = ParagraphStyle(
            'SlideTitle',
            parent=styles['Heading1'],
            fontSize=20,
            spaceAfter=16,
            spaceBefore=8,
            alignment=TA_CENTER,
            textColor=colors.darkblue,
            fontName='Helvetica-Bold',
            leading=24  # Improved line spacing
        )
        
        content_title_style = ParagraphStyle(
            'ContentTitle',
            parent=styles['Heading2'],
            fontSize=16,
            spaceAfter=10,
            spaceBefore=8,
            alignment=TA_LEFT,
            textColor=colors.darkblue,
            fontName='Helvetica-Bold',
            leading=20  # Improved line spacing
        )
        
        content_style = ParagraphStyle(
            'SlideContent',
            parent=styles['Normal'],
            fontSize=12,
            spaceAfter=10,
            spaceBefore=4,
            alignment=TA_LEFT,
            leading=16,  # Improved line spacing
            fontName='Helvetica'
        )
        
        bullet_style = ParagraphStyle(
            'BulletStyle',
            parent=content_style,
            leftIndent=30,
            bulletIndent=15,
            spaceAfter=8,
            bulletFontName='Symbol',
            bulletText='•',
            leading=16  # Improved line spacing
        )
        
        # Try to extract presentation title from metadata or first slide
        presentation_title = "PowerPoint Presentation"
        try:
            if hasattr(prs, 'core_properties') and prs.core_properties.title:
                presentation_title = prs.core_properties.title
            elif prs.slides and hasattr(prs.slides[0], 'shapes'):
                for shape in prs.slides[0].shapes:
                    if hasattr(shape, 'text') and shape.text.strip() and len(shape.text) < 100:
                        presentation_title = shape.text.strip()
                        break
        except:
            pass
        
        # Add presentation title
        story.append(Paragraph(presentation_title, slide_title_style))
        story.append(Spacer(1, 20))
        
        # Limit number of slides to prevent excessive processing
        max_slides = 50  # Increased limit for better content coverage
        slides_to_process = list(prs.slides)[:max_slides]
        
        for i, slide in enumerate(slides_to_process):
            # Add slide header with better formatting
            slide_header = f"Slide {i + 1} of {len(slides_to_process)}"
            story.append(Paragraph(slide_header, slide_title_style))
            story.append(Spacer(1, 16))
            
            # Process slide content with increased limits
            shape_count = 0
            max_shapes_per_slide = 30  # Increased limit for better content capture
            slide_has_content = False
            
            # Try to extract slide title first
            slide_title = None
            try:
                # Look for title placeholder
                for shape in slide.shapes:
                    if hasattr(shape, 'is_placeholder') and shape.is_placeholder and shape.placeholder_format.idx == 0:
                        if hasattr(shape, 'text') and shape.text.strip():
                            slide_title = shape.text.strip()
                            break
                    # Fallback: look for any text that looks like a title
                    elif hasattr(shape, 'text') and shape.text.strip() and len(shape.text) < 100:
                        if shape.text.strip() not in slide_header:  # Avoid duplicating slide number
                            slide_title = shape.text.strip()
                            break
            except:
                pass
            
            # Add slide title if found
            if slide_title:
                story.append(Paragraph(slide_title, content_title_style))
                story.append(Spacer(1, 12))
                slide_has_content = True
            
            # Sort shapes by their position (top to bottom, left to right)
            try:
                sorted_shapes = sorted(slide.shapes, key=lambda s: (s.top, s.left) if hasattr(s, 'top') and hasattr(s, 'left') else (0, 0))
            except:
                sorted_shapes = slide.shapes
            
            for shape in sorted_shapes:
                # Skip if this is the title we already processed
                if slide_title and hasattr(shape, 'text') and shape.text.strip() == slide_title:
                    continue
                    
                shape_count += 1
                if shape_count > max_shapes_per_slide:
                    break
                
                try:
                    # Handle images with improved quality
                    if hasattr(shape, 'shape_type') and shape.shape_type == 13:  # Picture type
                        try:
                            # Extract image from shape
                            image = shape.image
                            image_bytes = image.blob
                            
                            # Create PIL Image
                            img_buffer = io.BytesIO(image_bytes)
                            pil_image = Image.open(img_buffer)
                            
                            # Calculate better image size based on PDF dimensions
                            max_width = min(pdf_doc.width * 0.8, 500)  # 80% of page width or 500pt max
                            max_height = min(pdf_doc.height * 0.6, 400)  # 60% of page height or 400pt max
                            
                            # Preserve aspect ratio
                            width, height = pil_image.size
                            aspect = width / height if height > 0 else 1
                            
                            if width > max_width:
                                width = max_width
                                height = width / aspect
                            
                            if height > max_height:
                                height = max_height
                                width = height * aspect
                            
                            # Ensure dimensions are valid
                            width = max(1, int(width))
                            height = max(1, int(height))
                            
                            # Resize image with high quality
                            try:
                                pil_image = pil_image.resize((width, height), Image.Resampling.LANCZOS)
                            except:
                                # Fallback to BICUBIC if LANCZOS fails
                                pil_image = pil_image.resize((width, height), Image.BICUBIC)
                            
                            # Convert to RGB if necessary
                            if pil_image.mode in ('RGBA', 'LA', 'P'):
                                pil_image = pil_image.convert('RGB')
                            
                            # Save to BytesIO with higher quality
                            img_buffer = io.BytesIO()
                            pil_image.save(img_buffer, format='JPEG', quality=95)
                            img_buffer.seek(0)
                            
                            # Create ReportLab image with proper alignment
                            rl_image = RLImage(img_buffer, width=width, height=height)
                            story.append(rl_image)
                            story.append(Spacer(1, 16))
                            slide_has_content = True
                            
                        except Exception as img_error:
                            # If image extraction fails, add a placeholder
                            story.append(Paragraph("[Image could not be extracted]", content_style))
                            story.append(Spacer(1, 8))
                            slide_has_content = True
                    
                    # Handle text content with improved formatting
                    elif hasattr(shape, 'text') and shape.text.strip():
                        text = shape.text.strip()
                        if len(text) > 2000:  # Increased text limit
                            text = text[:2000] + "..."
                        
                        # Escape special characters for ReportLab
                        text = text.replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;')
                        
                        # Determine text type and apply appropriate styling
                        if len(text) < 100 and '\n' not in text:
                            # Likely a title or header
                            current_style = content_title_style
                        elif text.startswith(('•', '-', '*', '◦', '▪', '▫')):
                            # Bullet point
                            current_style = bullet_style
                            text = text[1:].strip()  # Remove bullet character
                        else:
                            current_style = content_style
                        
                        # Handle multi-line text better
                        paragraphs = text.split('\n')
                        for para in paragraphs:
                            if para.strip():
                                story.append(Paragraph(para.strip(), current_style))
                                story.append(Spacer(1, 4))
                        
                        story.append(Spacer(1, 8))
                        slide_has_content = True
                    
                    # Enhanced text frame processing with better formatting
                    elif hasattr(shape, 'text_frame') and shape.text_frame:
                        text_frame = shape.text_frame
                        if hasattr(text_frame, 'paragraphs'):
                            for paragraph in text_frame.paragraphs:
                                if paragraph.text.strip():
                                    para_text = paragraph.text.strip()
                                    if len(para_text) > 1500:
                                        para_text = para_text[:1500] + "..."
                                    
                                    # Escape special characters for ReportLab
                                    para_text = para_text.replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;')
                                    
                                    # Enhanced bullet detection and formatting
                                    level = getattr(paragraph, 'level', 0)
                                    
                                    if (para_text.startswith(('•', '-', '*', '◦', '▪', '▫')) or level > 0):
                                        # Create custom bullet style based on level
                                        level_style = ParagraphStyle(
                                            f'BulletLevel{level}',
                                            parent=bullet_style,
                                            leftIndent=30 + (level * 15),
                                            bulletIndent=15 + (level * 15)
                                        )
                                        current_style = level_style
                                        
                                        if para_text.startswith(('•', '-', '*', '◦', '▪', '▫')):
                                            para_text = para_text[1:].strip()
                                    elif len(para_text) < 100 and '\n' not in para_text:
                                        current_style = content_title_style
                                    else:
                                        current_style = content_style
                                    
                                    story.append(Paragraph(para_text, current_style))
                                    story.append(Spacer(1, 4))
                                    slide_has_content = True
                    
                    # Enhanced table handling with better formatting
                    elif hasattr(shape, 'table'):
                        table = shape.table
                        table_data = []
                        
                        # Process table data with better formatting
                        for row_idx, row in enumerate(table.rows):
                            row_data = []
                            for cell in row.cells:
                                cell_text = cell.text.strip() if cell.text else ""
                                if len(cell_text) > 200:  # Increased cell text limit
                                    cell_text = cell_text[:200] + "..."
                                
                                # Escape special characters for ReportLab
                                cell_text = cell_text.replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;')
                                row_data.append(cell_text)
                            
                            table_data.append(row_data)
                        
                        if table_data and any(any(cell for cell in row) for row in table_data):
                            # Calculate better column widths
                            col_count = len(table_data[0]) if table_data else 0
                            if col_count > 0:
                                available_width = pdf_doc.width * 0.9  # 90% of page width
                                col_widths = [available_width / col_count] * col_count
                                
                                # Create enhanced ReportLab table with better styling
                                t = Table(table_data, repeatRows=1, colWidths=col_widths)
                                t.setStyle(TableStyle([
                                    ('BACKGROUND', (0, 0), (-1, 0), colors.lightblue),
                                    ('TEXTCOLOR', (0, 0), (-1, 0), colors.black),
                                    ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
                                    ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
                                    ('FONTNAME', (0, 1), (-1, -1), 'Helvetica'),
                                    ('FONTSIZE', (0, 0), (-1, -1), 10),  # Increased font size
                                    ('GRID', (0, 0), (-1, -1), 1, colors.black),
                                    ('VALIGN', (0, 0), (-1, -1), 'TOP'),
                                    ('ROWBACKGROUNDS', (0, 1), (-1, -1), [colors.white, colors.lightgrey]),
                                    ('LEFTPADDING', (0, 0), (-1, -1), 6),
                                    ('RIGHTPADDING', (0, 0), (-1, -1), 6),
                                    ('TOPPADDING', (0, 0), (-1, -1), 4),
                                    ('BOTTOMPADDING', (0, 0), (-1, -1), 4),
                                ]))
                                story.append(t)
                                story.append(Spacer(1, 16))
                                slide_has_content = True
                
                except Exception as shape_error:
                    # Skip problematic shapes but continue processing
                    continue
            
            # Add content if slide was empty
            if not slide_has_content:
                story.append(Paragraph("(Empty slide or content could not be extracted)", content_style))
                story.append(Spacer(1, 12))
            
            # Add separator between slides (except for last slide)
            if i < len(slides_to_process) - 1:
                story.append(Spacer(1, 20))
                story.append(PageBreak())
        
        # Safety check: limit total story elements
        if len(story) > 500:  # Increased limit for better content
            story = story[:500]
            story.append(Paragraph("... (Content truncated for safety - presentation too large)", content_style))
        
        # Build PDF with enhanced error handling
        try:
            pdf_doc.build(story)
        except Exception as build_error:
            # Enhanced fallback: create better formatted simple PDF
            st.warning("Complex layout failed, creating simplified PDF...")
            
            # Reset output buffer
            output.seek(0)
            output.truncate(0)
            
            # Use letter size for fallback
            pdf_doc = SimpleDocTemplate(
                output,
                pagesize=letter,
                leftMargin=0.75*inch,
                rightMargin=0.75*inch,
                topMargin=0.75*inch,
                bottomMargin=0.75*inch
            )
            
            simple_story = []
            simple_story.append(Paragraph(presentation_title, styles['Title']))
            simple_story.append(Spacer(1, 20))
            
            # Extract all content with better formatting
            for i, slide in enumerate(slides_to_process[:20]):  # Increased limit to 20 slides for fallback
                simple_story.append(Paragraph(f"Slide {i + 1}", styles['Heading2']))
                simple_story.append(Spacer(1, 12))
                
                # Try to extract slide title
                slide_title = None
                try:
                    for shape in slide.shapes:
                        if hasattr(shape, 'is_placeholder') and shape.is_placeholder and shape.placeholder_format.idx == 0:
                            if hasattr(shape, 'text') and shape.text.strip():
                                slide_title = shape.text.strip()
                                break
                except:
                    pass
                
                # Add slide title if found
                if slide_title:
                    simple_story.append(Paragraph(slide_title, styles['Heading3']))
                    simple_story.append(Spacer(1, 8))
                
                # Process shapes with better formatting
                for shape in slide.shapes:
                    try:
                        # Skip if this is the title we already processed
                        if slide_title and hasattr(shape, 'text') and shape.text.strip() == slide_title:
                            continue
                            
                        if hasattr(shape, 'text') and shape.text.strip():
                            text = shape.text.strip()[:1000]  # Increased text length
                            
                            # Escape special characters for ReportLab
                            text = text.replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;')
                            
                            # Handle multi-line text better
                            paragraphs = text.split('\n')
                            for para in paragraphs:
                                if para.strip():
                                    simple_story.append(Paragraph(para.strip(), styles['Normal']))
                                    simple_story.append(Spacer(1, 4))
                            
                            simple_story.append(Spacer(1, 8))
                        elif hasattr(shape, 'shape_type') and shape.shape_type == 13:
                            simple_story.append(Paragraph("[Image present but not extracted in simplified mode]", styles['Normal']))
                            simple_story.append(Spacer(1, 6))
                    except:
                        continue
                
                simple_story.append(Spacer(1, 16))
                
                # Add page break between slides
                if i < len(slides_to_process[:20]) - 1:
                    simple_story.append(PageBreak())
            
            # Build simple version
            pdf_doc.build(simple_story)
        
        output.seek(0)
        return output
    except Exception as e:
        st.error(f"PowerPoint conversion error: {str(e)}")
        return None

def image_to_pdf(uploaded_file):
    """Convert Image to PDF"""
    try:
        img = Image.open(uploaded_file)
        
        if img.mode == 'RGBA':
            img = img.convert('RGB')
        
        output = io.BytesIO()
        img_byte_arr = io.BytesIO()
        img.save(img_byte_arr, format='JPEG')
        img_byte_arr.seek(0)
        
        pdf_bytes = img2pdf.convert(img_byte_arr.getvalue())
        output.write(pdf_bytes)
        output.seek(0)
        return output
    except Exception as e:
        st.error(f"Error: {str(e)}")
        return None

def text_to_pdf(text_content):
    """Convert Text to PDF"""
    try:
        output = io.BytesIO()
        pdf_doc = SimpleDocTemplate(output, pagesize=letter)
        styles = getSampleStyleSheet()
        story = []
        
        for line in text_content.split('\n'):
            if line.strip():
                story.append(Paragraph(line, styles['Normal']))
                story.append(Spacer(1, 12))
        
        pdf_doc.build(story)
        output.seek(0)
        return output
    except Exception as e:
        st.error(f"Error: {str(e)}")
        return None

def pdf_to_word(uploaded_file):
    """Convert PDF to Word"""
    try:
        pdf_reader = PyPDF2.PdfReader(uploaded_file)
        doc = Document()
        
        for page in pdf_reader.pages:
            text = page.extract_text()
            doc.add_paragraph(text)
        
        output = io.BytesIO()
        doc.save(output)
        output.seek(0)
        return output
    except Exception as e:
        st.error(f"Error: {str(e)}")
        return None

def pdf_to_excel(uploaded_file):
    """Convert PDF to Excel"""
    try:
        pdf_reader = PyPDF2.PdfReader(uploaded_file)
        data = []
        
        for page in pdf_reader.pages:
            text = page.extract_text()
            lines = text.split('\n')
            for line in lines:
                if line.strip():
                    data.append([line.strip()])
        
        df = pd.DataFrame(data, columns=['Content'])
        output = io.BytesIO()
        df.to_excel(output, index=False, engine='openpyxl')
        output.seek(0)
        return output
    except Exception as e:
        st.error(f"Error: {str(e)}")
        return None

def pdf_to_ppt(uploaded_file):
    """Convert PDF to PowerPoint"""
    try:
        pdf_reader = PyPDF2.PdfReader(uploaded_file)
        prs = Presentation()
        
        for page_num, page in enumerate(pdf_reader.pages):
            text = page.extract_text()
            slide = prs.slides.add_slide(prs.slide_layouts[1])
            title = slide.shapes.title
            title.text = f"Page {page_num + 1}"
            
            body = slide.placeholders[1]
            tf = body.text_frame
            tf.text = text
        
        output = io.BytesIO()
        prs.save(output)
        output.seek(0)
        return output
    except Exception as e:
        st.error(f"Error: {str(e)}")
        return None

def pdf_to_text(uploaded_file):
    """Convert PDF to Text"""
    try:
        pdf_reader = PyPDF2.PdfReader(uploaded_file)
        text = ""
        
        for page in pdf_reader.pages:
            text += page.extract_text() + "\n\n"
        
        return text
    except Exception as e:
        st.error(f"Error: {str(e)}")
        return None

def pdf_to_images(uploaded_file, format='JPEG'):
    """Convert PDF pages to images"""
    try:
        images = convert_from_bytes(uploaded_file.read(), dpi=200)
        
        zip_buffer = io.BytesIO()
        with zipfile.ZipFile(zip_buffer, 'w', zipfile.ZIP_DEFLATED) as zip_file:
            for i, image in enumerate(images):
                img_byte_arr = io.BytesIO()
                if format == 'JPEG' and image.mode == 'RGBA':
                    image = image.convert('RGB')
                image.save(img_byte_arr, format=format)
                img_byte_arr.seek(0)
                zip_file.writestr(f'page_{i+1}.{format.lower()}', img_byte_arr.getvalue())
        
        zip_buffer.seek(0)
        return zip_buffer
    except Exception as e:
        st.error(f"Error: {str(e)}\nNote: Make sure poppler-utils is installed for PDF to image conversion")
        return None

def merge_pdfs(uploaded_files):
    """Merge multiple PDFs with memory optimization"""
    try:
        # Validate total file size
        total_size = sum(len(file.getvalue()) for file in uploaded_files)
        if total_size > 100 * 1024 * 1024:  # 100MB limit
            raise ValueError("Total file size too large. Please keep total size under 100MB.")
        
        pdf_writer = PyPDF2.PdfWriter()
        total_pages = 0
        
        for uploaded_file in uploaded_files:
            # Reset file pointer
            uploaded_file.seek(0)
            
            # Read PDF
            pdf_reader = PyPDF2.PdfReader(uploaded_file)
            current_pages = len(pdf_reader.pages)
            
            # Check page limit
            total_pages += current_pages
            if total_pages > 500:  # Limit total pages
                raise ValueError("Too many pages. Please keep total pages under 500.")
            
            # Add pages with memory optimization
            for page in pdf_reader.pages:
                pdf_writer.add_page(page)
                # Clear page from memory
                page.clear()
        
        # Write output with compression
        output = io.BytesIO()
        pdf_writer.write(output)
        output.seek(0)
        
        # Clear writer from memory
        pdf_writer.close()
        return output
    except Exception as e:
        st.error(f"Error: {str(e)}")
        return None

def split_pdf(uploaded_file, split_at):
    """Split PDF at specific page"""
    try:
        pdf_reader = PyPDF2.PdfReader(uploaded_file)
        
        zip_buffer = io.BytesIO()
        with zipfile.ZipFile(zip_buffer, 'w', zipfile.ZIP_DEFLATED) as zip_file:
            pdf_writer1 = PyPDF2.PdfWriter()
            for i in range(min(split_at, len(pdf_reader.pages))):
                pdf_writer1.add_page(pdf_reader.pages[i])
            
            output1 = io.BytesIO()
            pdf_writer1.write(output1)
            output1.seek(0)
            zip_file.writestr('part1.pdf', output1.getvalue())
            
            if split_at < len(pdf_reader.pages):
                pdf_writer2 = PyPDF2.PdfWriter()
                for i in range(split_at, len(pdf_reader.pages)):
                    pdf_writer2.add_page(pdf_reader.pages[i])
                
                output2 = io.BytesIO()
                pdf_writer2.write(output2)
                output2.seek(0)
                zip_file.writestr('part2.pdf', output2.getvalue())
        
        zip_buffer.seek(0)
        return zip_buffer
    except Exception as e:
        st.error(f"Error: {str(e)}")
        return None

def compress_pdf(uploaded_file):
    """Compress PDF (basic implementation)"""
    try:
        pdf_reader = PyPDF2.PdfReader(uploaded_file)
        pdf_writer = PyPDF2.PdfWriter()
        
        for page in pdf_reader.pages:
            page.compress_content_streams()
            pdf_writer.add_page(page)
        
        output = io.BytesIO()
        pdf_writer.write(output)
        output.seek(0)
        return output
    except Exception as e:
        st.error(f"Error: {str(e)}")
        return None

def rotate_pdf(uploaded_file, rotation):
    """Rotate PDF pages"""
    try:
        pdf_reader = PyPDF2.PdfReader(uploaded_file)
        pdf_writer = PyPDF2.PdfWriter()
        
        for page in pdf_reader.pages:
            page.rotate(rotation)
            pdf_writer.add_page(page)
        
        output = io.BytesIO()
        pdf_writer.write(output)
        output.seek(0)
        return output
    except Exception as e:
        st.error(f"Error: {str(e)}")
        return None

def remove_pdf_pages(uploaded_file, pages_to_remove):
    """Remove specific pages from PDF"""
    try:
        pdf_reader = PyPDF2.PdfReader(uploaded_file)
        pdf_writer = PyPDF2.PdfWriter()
        
        for i, page in enumerate(pdf_reader.pages):
            if i + 1 not in pages_to_remove:
                pdf_writer.add_page(page)
        
        output = io.BytesIO()
        pdf_writer.write(output)
        output.seek(0)
        return output
    except Exception as e:
        st.error(f"Error: {str(e)}")
        return None

def extract_pdf_pages(uploaded_file, pages_to_extract):
    """Extract specific pages from PDF"""
    try:
        pdf_reader = PyPDF2.PdfReader(uploaded_file)
        pdf_writer = PyPDF2.PdfWriter()
        
        for page_num in pages_to_extract:
            if 0 < page_num <= len(pdf_reader.pages):
                pdf_writer.add_page(pdf_reader.pages[page_num - 1])
        
        output = io.BytesIO()
        pdf_writer.write(output)
        output.seek(0)
        return output
    except Exception as e:
        st.error(f"Error: {str(e)}")
        return None

def convert_image_format(uploaded_file, output_format):
    """Convert between image formats"""
    try:
        img = Image.open(uploaded_file)
        
        if output_format in ['JPEG', 'JPG'] and img.mode in ('RGBA', 'P'):
            img = img.convert('RGB')
        
        output = io.BytesIO()
        save_format = 'JPEG' if output_format == 'JPG' else output_format
        img.save(output, format=save_format)
        output.seek(0)
        return output
    except Exception as e:
        st.error(f"Error: {str(e)}")
        return None

def resize_image(uploaded_file, width, height, maintain_aspect=True):
    """Resize image"""
    try:
        img = Image.open(uploaded_file)
        
        if maintain_aspect:
            img.thumbnail((width, height), Image.Resampling.LANCZOS)
        else:
            img = img.resize((width, height), Image.Resampling.LANCZOS)
        
        output = io.BytesIO()
        img_format = img.format or 'PNG'
        if img_format == 'JPEG' and img.mode in ('RGBA', 'P'):
            img = img.convert('RGB')
        img.save(output, format=img_format)
        output.seek(0)
        return output
    except Exception as e:
        st.error(f"Error: {str(e)}")
        return None

def rotate_image(uploaded_file, angle):
    """Rotate image"""
    try:
        img = Image.open(uploaded_file)
        rotated = img.rotate(angle, expand=True)
        
        output = io.BytesIO()
        img_format = img.format or 'PNG'
        rotated.save(output, format=img_format)
        output.seek(0)
        return output
    except Exception as e:
        st.error(f"Error: {str(e)}")
        return None

def word_to_excel(uploaded_file):
    """Convert Word tables to Excel"""
    try:
        doc = Document(uploaded_file)
        
        all_data = []
        for table in doc.tables:
            for row in table.rows:
                row_data = [cell.text for cell in row.cells]
                all_data.append(row_data)
            all_data.append([])  # Empty row between tables
        
        if not all_data:
            st.warning("No tables found in document. Extracting text...")
            all_data = [[para.text] for para in doc.paragraphs if para.text.strip()]
        
        df = pd.DataFrame(all_data)
        output = io.BytesIO()
        df.to_excel(output, index=False, header=False, engine='openpyxl')
        output.seek(0)
        return output
    except Exception as e:
        st.error(f"Error: {str(e)}")
        return None

def excel_to_word(uploaded_file):
    """Convert Excel to Word"""
    try:
        df = pd.read_excel(uploaded_file)
        doc = Document()
        
        doc.add_heading('Excel Data', 0)
        
        table = doc.add_table(rows=len(df) + 1, cols=len(df.columns))
        table.style = 'Light Grid Accent 1'
        
        for i, column in enumerate(df.columns):
            table.rows[0].cells[i].text = str(column)
        
        for i, row in enumerate(df.itertuples(index=False), 1):
            for j, value in enumerate(row):
                table.rows[i].cells[j].text = str(value)
        
        output = io.BytesIO()
        doc.save(output)
        output.seek(0)
        return output
    except Exception as e:
        st.error(f"Error: {str(e)}")
        return None

def json_to_excel(json_data):
    """Convert JSON to Excel"""
    try:
        data = json.loads(json_data)
        
        if isinstance(data, list):
            df = pd.DataFrame(data)
        elif isinstance(data, dict):
            df = pd.DataFrame([data])
        else:
            st.error("Invalid JSON format")
            return None
        
        output = io.BytesIO()
        df.to_excel(output, index=False, engine='openpyxl')
        output.seek(0)
        return output
    except Exception as e:
        st.error(f"Error: {str(e)}")
        return None

def excel_to_json(uploaded_file):
    """Convert Excel to JSON"""
    try:
        df = pd.read_excel(uploaded_file)
        json_data = df.to_json(orient='records', indent=2)
        return json_data
    except Exception as e:
        st.error(f"Error: {str(e)}")
        return None

# Main conversion interface
st.markdown(f"### {conversion_type}")

# Conversion logic
if conversion_type == "Word to PDF":
    uploaded_file = st.file_uploader("Upload Word Document", type=['docx', 'doc'])
    if uploaded_file and st.button("Convert to PDF"):
        with st.spinner("Converting..."):
            result = word_to_pdf(uploaded_file)
            if result:
                st.success("✅ Conversion successful!")
                st.download_button("📥 Download PDF", result, f"{Path(uploaded_file.name).stem}.pdf", "application/pdf")

elif conversion_type == "Excel to PDF":
    uploaded_file = st.file_uploader("Upload Excel File", type=['xlsx', 'xls'])
    if uploaded_file and st.button("Convert to PDF"):
        with st.spinner("Converting..."):
            result = excel_to_pdf(uploaded_file)
            if result:
                st.success("✅ Conversion successful!")
                st.download_button("📥 Download PDF", result, f"{Path(uploaded_file.name).stem}.pdf", "application/pdf")

elif conversion_type == "PowerPoint to PDF":
    uploaded_file = st.file_uploader("Upload PowerPoint", type=['pptx', 'ppt'])
    if uploaded_file and st.button("Convert to PDF"):
        with st.spinner("Converting..."):
            result = ppt_to_pdf(uploaded_file)
            if result:
                st.success("✅ Conversion successful!")
                st.download_button("📥 Download PDF", result, f"{Path(uploaded_file.name).stem}.pdf", "application/pdf")

elif conversion_type in ["JPG to PDF", "PNG to PDF"]:
    file_type = conversion_type.split()[0].lower()
    uploaded_file = st.file_uploader(f"Upload {file_type.upper()} Image", type=[file_type, 'jpeg'])
    if uploaded_file and st.button("Convert to PDF"):
        with st.spinner("Converting..."):
            result = image_to_pdf(uploaded_file)
            if result:
                st.success("✅ Conversion successful!")
                st.download_button("📥 Download PDF", result, f"{Path(uploaded_file.name).stem}.pdf", "application/pdf")

elif conversion_type == "Text to PDF":
    text_input = st.text_area("Enter text to convert", height=300)
    if text_input and st.button("Convert to PDF"):
        with st.spinner("Converting..."):
            result = text_to_pdf(text_input)
            if result:
                st.success("✅ Conversion successful!")
                st.download_button("📥 Download PDF", result, "text_document.pdf", "application/pdf")

elif conversion_type == "PDF to Word":
    uploaded_file = st.file_uploader("Upload PDF", type=['pdf'])
    if uploaded_file and st.button("Convert to Word"):
        with st.spinner("Converting..."):
            result = pdf_to_word(uploaded_file)
            if result:
                st.success("✅ Conversion successful!")
                st.download_button("📥 Download Word", result, f"{Path(uploaded_file.name).stem}.docx", 
                                 "application/vnd.openxmlformats-officedocument.wordprocessingml.document")

elif conversion_type == "PDF to Excel":
    uploaded_file = st.file_uploader("Upload PDF", type=['pdf'])
    if uploaded_file and st.button("Convert to Excel"):
        with st.spinner("Converting..."):
            result = pdf_to_excel(uploaded_file)
            if result:
                st.success("✅ Conversion successful!")
                st.download_button("📥 Download Excel", result, f"{Path(uploaded_file.name).stem}.xlsx",
                                 "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

elif conversion_type == "PDF to PowerPoint":
    uploaded_file = st.file_uploader("Upload PDF", type=['pdf'])
    if uploaded_file and st.button("Convert to PowerPoint"):
        with st.spinner("Converting..."):
            result = pdf_to_ppt(uploaded_file)
            if result:
                st.success("✅ Conversion successful!")
                st.download_button("📥 Download PowerPoint", result, f"{Path(uploaded_file.name).stem}.pptx",
                                 "application/vnd.openxmlformats-officedocument.presentationml.presentation")

elif conversion_type == "PDF to Text":
    uploaded_file = st.file_uploader("Upload PDF", type=['pdf'])
    if uploaded_file and st.button("Extract Text"):
        with st.spinner("Extracting..."):
            result = pdf_to_text(uploaded_file)
            if result:
                st.success("✅ Text extracted successfully!")
                st.text_area("Extracted Text", result, height=300)
                st.download_button("📥 Download Text", result, f"{Path(uploaded_file.name).stem}.txt", "text/plain")

elif conversion_type in ["PDF to JPG", "PDF to PNG"]:
    format_type = conversion_type.split()[-1].upper()
    uploaded_file = st.file_uploader("Upload PDF", type=['pdf'])
    if uploaded_file and st.button(f"Convert to {format_type}"):
        with st.spinner("Converting..."):
            result = pdf_to_images(uploaded_file, format_type)
            if result:
                st.success("✅ Conversion successful! Multiple images will be downloaded as ZIP")
                st.download_button("📥 Download Images (ZIP)", result, f"{Path(uploaded_file.name).stem}_images.zip", "application/zip")

elif conversion_type == "Extract PDF Images":
    uploaded_file = st.file_uploader("Upload PDF", type=['pdf'])
    if uploaded_file and st.button("Extract Images"):
        with st.spinner("Extracting..."):
            result = pdf_to_images(uploaded_file, 'PNG')
            if result:
                st.success("✅ Images extracted successfully!")
                st.download_button("📥 Download Images (ZIP)", result, f"{Path(uploaded_file.name).stem}_extracted.zip", "application/zip")

elif conversion_type == "Merge PDF":
    uploaded_files = st.file_uploader("Upload PDF files to merge", type=['pdf'], accept_multiple_files=True)
    if uploaded_files and len(uploaded_files) > 1 and st.button("Merge PDFs"):
        with st.spinner("Merging..."):
            result = merge_pdfs(uploaded_files)
            if result:
                st.success("✅ PDFs merged successfully!")
                st.download_button("📥 Download Merged PDF", result, "merged.pdf", "application/pdf")
    elif uploaded_files and len(uploaded_files) == 1:
        st.warning("Please upload at least 2 PDF files to merge")

elif conversion_type == "Split PDF":
    uploaded_file = st.file_uploader("Upload PDF", type=['pdf'])
    if uploaded_file:
        pdf_reader = PyPDF2.PdfReader(uploaded_file)
        total_pages = len(pdf_reader.pages)
        st.info(f"📄 PDF has {total_pages} pages")
        
        split_at = st.number_input("Split at page number", min_value=1, max_value=total_pages-1, value=1)
        
        if st.button("Split PDF"):
            with st.spinner("Splitting..."):
                result = split_pdf(uploaded_file, split_at)
                if result:
                    st.success("✅ PDF split successfully!")
                    st.download_button("📥 Download Split PDFs (ZIP)", result, f"{Path(uploaded_file.name).stem}_split.zip", "application/zip")

elif conversion_type == "Compress PDF":
    uploaded_file = st.file_uploader("Upload PDF", type=['pdf'])
    if uploaded_file and st.button("Compress PDF"):
        with st.spinner("Compressing..."):
            result = compress_pdf(uploaded_file)
            if result:
                original_size = len(uploaded_file.getvalue())
                compressed_size = len(result.getvalue())
                reduction = ((original_size - compressed_size) / original_size) * 100
                
                st.success(f"✅ PDF compressed successfully!")
                st.info(f"Size reduction: {reduction:.1f}%")
                st.download_button("📥 Download Compressed PDF", result, f"{Path(uploaded_file.name).stem}_compressed.pdf", "application/pdf")

elif conversion_type == "Rotate PDF":
    uploaded_file = st.file_uploader("Upload PDF", type=['pdf'])
    if uploaded_file:
        rotation = st.selectbox("Select rotation angle", [90, 180, 270])
        
        if st.button("Rotate PDF"):
            with st.spinner("Rotating..."):
                result = rotate_pdf(uploaded_file, rotation)
                if result:
                    st.success("✅ PDF rotated successfully!")
                    st.download_button("📥 Download Rotated PDF", result, f"{Path(uploaded_file.name).stem}_rotated.pdf", "application/pdf")

elif conversion_type == "Remove PDF Pages":
    uploaded_file = st.file_uploader("Upload PDF", type=['pdf'])
    if uploaded_file:
        pdf_reader = PyPDF2.PdfReader(uploaded_file)
        total_pages = len(pdf_reader.pages)
        st.info(f"📄 PDF has {total_pages} pages")
        
        pages_input = st.text_input("Enter page numbers to remove (comma-separated, e.g., 1,3,5)")
        
        if st.button("Remove Pages"):
            try:
                pages_to_remove = [int(p.strip()) for p in pages_input.split(',') if p.strip()]
                with st.spinner("Removing pages..."):
                    result = remove_pdf_pages(uploaded_file, pages_to_remove)
                    if result:
                        st.success(f"✅ Removed {len(pages_to_remove)} pages successfully!")
                        st.download_button("📥 Download PDF", result, f"{Path(uploaded_file.name).stem}_modified.pdf", "application/pdf")
            except ValueError:
                st.error("Please enter valid page numbers")

elif conversion_type == "Extract PDF Pages":
    uploaded_file = st.file_uploader("Upload PDF", type=['pdf'])
    if uploaded_file:
        pdf_reader = PyPDF2.PdfReader(uploaded_file)
        total_pages = len(pdf_reader.pages)
        st.info(f"📄 PDF has {total_pages} pages")
        
        pages_input = st.text_input("Enter page numbers to extract (comma-separated, e.g., 1,3,5)")
        
        if st.button("Extract Pages"):
            try:
                pages_to_extract = [int(p.strip()) for p in pages_input.split(',') if p.strip()]
                with st.spinner("Extracting pages..."):
                    result = extract_pdf_pages(uploaded_file, pages_to_extract)
                    if result:
                        st.success(f"✅ Extracted {len(pages_to_extract)} pages successfully!")
                        st.download_button("📥 Download PDF", result, f"{Path(uploaded_file.name).stem}_extracted.pdf", "application/pdf")
            except ValueError:
                st.error("Please enter valid page numbers")

elif conversion_type in ["JPG to PNG", "PNG to JPG"]:
    input_format = conversion_type.split()[0].lower()
    output_format = conversion_type.split()[-1].upper()
    
    uploaded_file = st.file_uploader(f"Upload {input_format.upper()} Image", type=[input_format, 'jpeg'])
    if uploaded_file and st.button(f"Convert to {output_format}"):
        with st.spinner("Converting..."):
            result = convert_image_format(uploaded_file, output_format)
            if result:
                st.success("✅ Conversion successful!")
                ext = output_format.lower()
                st.download_button("📥 Download Image", result, f"{Path(uploaded_file.name).stem}.{ext}", f"image/{ext}")

elif conversion_type == "Image to WebP":
    uploaded_file = st.file_uploader("Upload Image", type=['jpg', 'jpeg', 'png', 'bmp'])
    if uploaded_file and st.button("Convert to WebP"):
        with st.spinner("Converting..."):
            result = convert_image_format(uploaded_file, 'WEBP')
            if result:
                st.success("✅ Conversion successful!")
                st.download_button("📥 Download WebP", result, f"{Path(uploaded_file.name).stem}.webp", "image/webp")

elif conversion_type == "WebP to JPG":
    uploaded_file = st.file_uploader("Upload WebP Image", type=['webp'])
    if uploaded_file and st.button("Convert to JPG"):
        with st.spinner("Converting..."):
            result = convert_image_format(uploaded_file, 'JPEG')
            if result:
                st.success("✅ Conversion successful!")
                st.download_button("📥 Download JPG", result, f"{Path(uploaded_file.name).stem}.jpg", "image/jpeg")

elif conversion_type == "WebP to PNG":
    uploaded_file = st.file_uploader("Upload WebP Image", type=['webp'])
    if uploaded_file and st.button("Convert to PNG"):
        with st.spinner("Converting..."):
            result = convert_image_format(uploaded_file, 'PNG')
            if result:
                st.success("✅ Conversion successful!")
                st.download_button("📥 Download PNG", result, f"{Path(uploaded_file.name).stem}.png", "image/png")

elif conversion_type == "Image to BMP":
    uploaded_file = st.file_uploader("Upload Image", type=['jpg', 'jpeg', 'png', 'webp'])
    if uploaded_file and st.button("Convert to BMP"):
        with st.spinner("Converting..."):
            result = convert_image_format(uploaded_file, 'BMP')
            if result:
                st.success("✅ Conversion successful!")
                st.download_button("📥 Download BMP", result, f"{Path(uploaded_file.name).stem}.bmp", "image/bmp")

elif conversion_type == "BMP to JPG":
    uploaded_file = st.file_uploader("Upload BMP Image", type=['bmp'])
    if uploaded_file and st.button("Convert to JPG"):
        with st.spinner("Converting..."):
            result = convert_image_format(uploaded_file, 'JPEG')
            if result:
                st.success("✅ Conversion successful!")
                st.download_button("📥 Download JPG", result, f"{Path(uploaded_file.name).stem}.jpg", "image/jpeg")

elif conversion_type == "Resize Image":
    uploaded_file = st.file_uploader("Upload Image", type=['jpg', 'jpeg', 'png', 'webp', 'bmp'])
    if uploaded_file:
        img = Image.open(uploaded_file)
        st.image(img, caption=f"Original: {img.size[0]}x{img.size[1]} pixels", use_container_width=True)
        
        col1, col2 = st.columns(2)
        with col1:
            width = st.number_input("Width (pixels)", min_value=1, value=img.size[0])
        with col2:
            height = st.number_input("Height (pixels)", min_value=1, value=img.size[1])
        
        maintain_aspect = st.checkbox("Maintain aspect ratio", value=True)
        
        if st.button("Resize Image"):
            with st.spinner("Resizing..."):
                result = resize_image(uploaded_file, width, height, maintain_aspect)
                if result:
                    st.success("✅ Image resized successfully!")
                    resized_img = Image.open(result)
                    st.image(resized_img, caption=f"Resized: {resized_img.size[0]}x{resized_img.size[1]} pixels")
                    result.seek(0)
                    ext = Path(uploaded_file.name).suffix
                    st.download_button("📥 Download Resized Image", result, f"{Path(uploaded_file.name).stem}_resized{ext}", f"image/{ext[1:]}")

elif conversion_type == "Rotate Image":
    uploaded_file = st.file_uploader("Upload Image", type=['jpg', 'jpeg', 'png', 'webp', 'bmp'])
    if uploaded_file:
        img = Image.open(uploaded_file)
        st.image(img, caption="Original Image", use_container_width=True)
        
        angle = st.selectbox("Select rotation angle", [90, 180, 270, -90, -180, -270])
        
        if st.button("Rotate Image"):
            with st.spinner("Rotating..."):
                result = rotate_image(uploaded_file, angle)
                if result:
                    st.success("✅ Image rotated successfully!")
                    rotated_img = Image.open(result)
                    st.image(rotated_img, caption=f"Rotated {angle}°")
                    result.seek(0)
                    ext = Path(uploaded_file.name).suffix
                    st.download_button("📥 Download Rotated Image", result, f"{Path(uploaded_file.name).stem}_rotated{ext}", f"image/{ext[1:]}")

elif conversion_type == "Word to Excel":
    uploaded_file = st.file_uploader("Upload Word Document", type=['docx'])
    if uploaded_file and st.button("Convert to Excel"):
        with st.spinner("Converting..."):
            result = word_to_excel(uploaded_file)
            if result:
                st.success("✅ Conversion successful!")
                st.download_button("📥 Download Excel", result, f"{Path(uploaded_file.name).stem}.xlsx",
                                 "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

elif conversion_type == "Excel to Word":
    uploaded_file = st.file_uploader("Upload Excel File", type=['xlsx', 'xls'])
    if uploaded_file and st.button("Convert to Word"):
        with st.spinner("Converting..."):
            result = excel_to_word(uploaded_file)
            if result:
                st.success("✅ Conversion successful!")
                st.download_button("📥 Download Word", result, f"{Path(uploaded_file.name).stem}.docx",
                                 "application/vnd.openxmlformats-officedocument.wordprocessingml.document")

elif conversion_type == "CSV to Excel":
    uploaded_file = st.file_uploader("Upload CSV", type=['csv'])
    if uploaded_file and st.button("Convert to Excel"):
        with st.spinner("Converting..."):
            df = pd.read_csv(uploaded_file)
            output = io.BytesIO()
            df.to_excel(output, index=False, engine='openpyxl')
            output.seek(0)
            st.success("✅ Conversion successful!")
            st.download_button("📥 Download Excel", output, f"{Path(uploaded_file.name).stem}.xlsx",
                             "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

elif conversion_type == "Excel to CSV":
    uploaded_file = st.file_uploader("Upload Excel", type=['xlsx', 'xls'])
    if uploaded_file and st.button("Convert to CSV"):
        with st.spinner("Converting..."):
            df = pd.read_excel(uploaded_file)
            csv_data = df.to_csv(index=False)
            st.success("✅ Conversion successful!")
            st.download_button("📥 Download CSV", csv_data, f"{Path(uploaded_file.name).stem}.csv", "text/csv")

elif conversion_type == "JSON to Excel":
    json_input = st.text_area("Paste JSON data", height=300)
    
    st.markdown("**Or upload a JSON file:**")
    uploaded_file = st.file_uploader("Upload JSON file", type=['json'])
    
    if uploaded_file:
        json_input = uploaded_file.read().decode('utf-8')
        st.text_area("JSON Content", json_input, height=200)
    
    if json_input and st.button("Convert to Excel"):
        with st.spinner("Converting..."):
            result = json_to_excel(json_input)
            if result:
                st.success("✅ Conversion successful!")
                st.download_button("📥 Download Excel", result, "converted.xlsx",
                                 "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

elif conversion_type == "Excel to JSON":
    uploaded_file = st.file_uploader("Upload Excel", type=['xlsx', 'xls'])
    if uploaded_file and st.button("Convert to JSON"):
        with st.spinner("Converting..."):
            result = excel_to_json(uploaded_file)
            if result:
                st.success("✅ Conversion successful!")
                st.text_area("JSON Output", result, height=300)
                st.download_button("📥 Download JSON", result, f"{Path(uploaded_file.name).stem}.json", "application/json")

# Footer
st.markdown("---")
st.markdown("### 📚 Supported Conversions")
col1, col2, col3 = st.columns(3)

with col1:
    st.markdown("**To PDF:**")
    st.markdown("- Word, Excel, PowerPoint")
    st.markdown("- JPG, PNG, Text")

with col2:
    st.markdown("**From PDF:**")
    st.markdown("- Word, Excel, PowerPoint")
    st.markdown("- JPG, PNG, Text")
    st.markdown("- Extract Images")

with col3:
    st.markdown("**Other:**")
    st.markdown("- Merge/Split PDF")
    st.markdown("- Image Formats & Resize")
    st.markdown("- CSV ↔ Excel ↔ JSON")

st.markdown("---")
st.markdown("**💡 Tips:**")
st.markdown("- For best results with PDF conversions, use clear, high-quality source files")
st.markdown("- Image to PDF conversions maintain original image quality")
st.markdown("- PDF to Image requires poppler-utils to be installed")
st.markdown("- Large files may take longer to process")