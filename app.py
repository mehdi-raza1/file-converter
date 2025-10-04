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
from reportlab.lib.pagesizes import letter
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Image as RLImage
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.lib.units import inch

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
        "AutoCAD to PDF",
        "Text to PDF"
    ],
    "📝 From PDF": [
        "PDF to Word",
        "PDF to Excel",
        "PDF to PowerPoint",
        "PDF to JPG",
        "PDF to PNG",
        "Extract PDF Images",
        "PDF to PDF/A"
    ],
    "🛠️ PDF Tools": [
        "Merge PDF",
        "Split PDF",
        "Compress PDF",
        "Rotate PDF"
    ],
    "🖼️ Image Conversion": [
        "JPG to PNG",
        "PNG to JPG",
        "Image to WebP",
        "WebP to JPG/PNG"
    ],
    "📊 Office Files": [
        "Word to Excel",
        "Excel to Word",
        "CSV to Excel",
        "Excel to CSV"
    ]
}

# Sidebar
st.sidebar.header("🎯 Select Conversion Type")
category = st.sidebar.selectbox("Category", list(conversion_categories.keys()))
conversion_type = st.sidebar.selectbox("Conversion", conversion_categories[category])

# Helper Functions

def word_to_pdf(uploaded_file):
    """Convert Word to PDF with formatting preserved using docx2pdf"""
    try:
        # Save uploaded file temporarily
        import tempfile
        import os
        from docx2pdf import convert
        
        with tempfile.NamedTemporaryFile(delete=False, suffix='.docx') as tmp_docx:
            tmp_docx.write(uploaded_file.read())
            tmp_docx_path = tmp_docx.name
        
        with tempfile.NamedTemporaryFile(delete=False, suffix='.pdf') as tmp_pdf:
            tmp_pdf_path = tmp_pdf.name
        
        # Convert with formatting preserved
        convert(tmp_docx_path, tmp_pdf_path)
        
        # Read the converted PDF
        with open(tmp_pdf_path, 'rb') as f:
            pdf_data = f.read()
        
        # Cleanup temporary files
        os.unlink(tmp_docx_path)
        os.unlink(tmp_pdf_path)
        
        output = io.BytesIO(pdf_data)
        return output
    except ImportError:
        st.warning("⚠️ docx2pdf not available. Using alternative method (formatting may not be perfect)...")
        return word_to_pdf_alternative(uploaded_file)
    except Exception as e:
        st.error(f"Error: {str(e)}")
        return None

def word_to_pdf_alternative(uploaded_file):
    """Alternative Word to PDF conversion with basic formatting"""
    try:
        from docx import Document
        from reportlab.lib.pagesizes import letter
        from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, Image as RLImage, PageBreak
        from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
        from reportlab.lib.enums import TA_LEFT, TA_CENTER, TA_RIGHT, TA_JUSTIFY
        from reportlab.lib import colors
        from reportlab.lib.units import inch
        
        doc = Document(uploaded_file)
        
        output = io.BytesIO()
        pdf_doc = SimpleDocTemplate(output, pagesize=letter,
                                   leftMargin=inch, rightMargin=inch,
                                   topMargin=inch, bottomMargin=inch)
        
        styles = getSampleStyleSheet()
        story = []
        
        # Create custom styles for different heading levels
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
        
        # Process paragraphs with formatting
        for para in doc.paragraphs:
            if not para.text.strip():
                story.append(Spacer(1, 6))
                continue
            
            # Detect heading styles
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
            
            # Handle text alignment
            if para.alignment == 1:  # Center
                style.alignment = TA_CENTER
            elif para.alignment == 2:  # Right
                style.alignment = TA_RIGHT
            elif para.alignment == 3:  # Justify
                style.alignment = TA_JUSTIFY
            
            # Build formatted text
            text = para.text
            
            # Handle bold and italic (basic)
            for run in para.runs:
                if run.bold and run.text:
                    text = text.replace(run.text, f"<b>{run.text}</b>")
                elif run.italic and run.text:
                    text = text.replace(run.text, f"<i>{run.text}</i>")
            
            p = Paragraph(text, style)
            story.append(p)
            story.append(Spacer(1, 6))
        
        # Process tables
        for table in doc.tables:
            data = []
            for row in table.rows:
                row_data = []
                for cell in row.cells:
                    row_data.append(cell.text)
                data.append(row_data)
            
            if data:
                t = Table(data)
                t.setStyle(TableStyle([
                    ('BACKGROUND', (0, 0), (-1, 0), colors.grey),
                    ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
                    ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
                    ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
                    ('FONTSIZE', (0, 0), (-1, 0), 12),
                    ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
                    ('BACKGROUND', (0, 1), (-1, -1), colors.beige),
                    ('GRID', (0, 0), (-1, -1), 1, colors.black)
                ]))
                story.append(t)
                story.append(Spacer(1, 12))
        
        pdf_doc.build(story)
        output.seek(0)
        return output
    except Exception as e:
        st.error(f"Error in alternative conversion: {str(e)}")
        return None

def excel_to_pdf(uploaded_file):
    """Convert Excel to PDF with formatting preserved"""
    try:
        import tempfile
        import os
        
        # Try using openpyxl with reportlab for better formatting
        wb = openpyxl.load_workbook(uploaded_file)
        
        output = io.BytesIO()
        pdf_doc = SimpleDocTemplate(output, pagesize=letter)
        styles = getSampleStyleSheet()
        story = []
        
        for sheet_name in wb.sheetnames:
            ws = wb[sheet_name]
            
            # Add sheet title
            story.append(Paragraph(f"<b>Sheet: {sheet_name}</b>", styles['Heading1']))
            story.append(Spacer(1, 12))
            
            # Extract data with formatting
            data = []
            for row in ws.iter_rows():
                row_data = []
                for cell in row:
                    value = cell.value if cell.value is not None else ""
                    row_data.append(str(value))
                data.append(row_data)
            
            if data:
                # Create table with styling
                from reportlab.platypus import Table, TableStyle
                from reportlab.lib import colors
                
                t = Table(data, repeatRows=1)
                
                # Apply styling
                table_style = [
                    ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#4472C4')),
                    ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
                    ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
                    ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
                    ('FONTSIZE', (0, 0), (-1, 0), 10),
                    ('BOTTOMPADDING', (0, 0), (-1, 0), 8),
                    ('TOPPADDING', (0, 0), (-1, 0), 8),
                    ('GRID', (0, 0), (-1, -1), 0.5, colors.grey),
                    ('ROWBACKGROUNDS', (0, 1), (-1, -1), [colors.white, colors.lightgrey])
                ]
                
                # Apply cell colors from Excel (if any)
                for row_idx, row in enumerate(ws.iter_rows(min_row=1)):
                    for col_idx, cell in enumerate(row):
                        if cell.fill and cell.fill.start_color:
                            try:
                                color_hex = cell.fill.start_color.rgb
                                if color_hex and len(color_hex) == 8:
                                    color = colors.HexColor(f"#{color_hex[2:]}")
                                    table_style.append(('BACKGROUND', (col_idx, row_idx), 
                                                      (col_idx, row_idx), color))
                            except:
                                pass
                
                t.setStyle(TableStyle(table_style))
                story.append(t)
                story.append(Spacer(1, 24))
        
        pdf_doc.build(story)
        output.seek(0)
        return output
    except Exception as e:
        st.error(f"Error: {str(e)}")
        return None

def ppt_to_pdf(uploaded_file):
    """Convert PowerPoint to PDF with better formatting"""
    try:
        from pptx import Presentation
        from reportlab.platypus import PageBreak, Image as RLImage
        from reportlab.lib.utils import ImageReader
        import tempfile
        
        prs = Presentation(uploaded_file)
        
        output = io.BytesIO()
        pdf_doc = SimpleDocTemplate(output, pagesize=letter)
        styles = getSampleStyleSheet()
        story = []
        
        # Custom style for slide titles
        from reportlab.lib.styles import ParagraphStyle
        from reportlab.lib.enums import TA_CENTER
        
        title_style = ParagraphStyle(
            'SlideTitle',
            parent=styles['Heading1'],
            fontSize=24,
            textColor=colors.HexColor('#4472C4'),
            spaceAfter=20,
            alignment=TA_CENTER
        )
        
        for i, slide in enumerate(prs.slides):
            # Slide number
            story.append(Paragraph(f"<b>Slide {i+1}</b>", title_style))
            story.append(Spacer(1, 12))
            
            # Process shapes
            for shape in slide.shapes:
                # Handle text boxes
                if hasattr(shape, "text") and shape.text:
                    text = shape.text.strip()
                    if text:
                        # Detect if it's a title
                        if shape.has_text_frame and len(shape.text_frame.paragraphs) > 0:
                            first_para = shape.text_frame.paragraphs[0]
                            if first_para.runs and len(first_para.runs) > 0:
                                run = first_para.runs[0]
                                
                                # Build formatted text
                                formatted_text = text
                                if run.font.bold:
                                    formatted_text = f"<b>{formatted_text}</b>"
                                if run.font.italic:
                                    formatted_text = f"<i>{formatted_text}</i>"
                                
                                # Choose style based on font size
                                if run.font.size and run.font.size > 300000:  # Large text
                                    para_style = styles['Heading2']
                                else:
                                    para_style = styles['Normal']
                                
                                story.append(Paragraph(formatted_text, para_style))
                            else:
                                story.append(Paragraph(text, styles['Normal']))
                        else:
                            story.append(Paragraph(text, styles['Normal']))
                        story.append(Spacer(1, 8))
                
                # Handle images
                if shape.shape_type == 13:  # Picture
                    try:
                        image = shape.image
                        image_bytes = image.blob
                        
                        img = ImageReader(io.BytesIO(image_bytes))
                        img_width, img_height = img.getSize()
                        
                        # Scale image to fit page
                        max_width = 6 * inch
                        max_height = 4 * inch
                        
                        aspect = img_height / float(img_width)
                        if img_width > max_width:
                            img_width = max_width
                            img_height = img_width * aspect
                        if img_height > max_height:
                            img_height = max_height
                            img_width = img_height / aspect
                        
                        rl_image = RLImage(io.BytesIO(image_bytes), width=img_width, height=img_height)
                        story.append(rl_image)
                        story.append(Spacer(1, 12))
                    except:
                        pass
                
                # Handle tables
                if shape.has_table:
                    table_data = []
                    for row in shape.table.rows:
                        row_data = []
                        for cell in row.cells:
                            row_data.append(cell.text)
                        table_data.append(row_data)
                    
                    if table_data:
                        t = Table(table_data)
                        t.setStyle(TableStyle([
                            ('BACKGROUND', (0, 0), (-1, 0), colors.grey),
                            ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
                            ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
                            ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
                            ('FONTSIZE', (0, 0), (-1, -1), 10),
                            ('GRID', (0, 0), (-1, -1), 1, colors.black),
                            ('ROWBACKGROUNDS', (0, 1), (-1, -1), [colors.white, colors.lightgrey])
                        ]))
                        story.append(t)
                        story.append(Spacer(1, 12))
            
            # Add page break between slides (except last)
            if i < len(prs.slides) - 1:
                story.append(PageBreak())
        
        pdf_doc.build(story)
        output.seek(0)
        return output
    except Exception as e:
        st.error(f"Error: {str(e)}")
        return None

def image_to_pdf(uploaded_file):
    """Convert Image to PDF"""
    try:
        img = Image.open(uploaded_file)
        
        # Convert to RGB if necessary
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
    """Convert PDF to Excel (basic text extraction)"""
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

def pdf_to_images(uploaded_file, format='JPEG'):
    """Convert PDF pages to images"""
    try:
        images = convert_from_bytes(uploaded_file.read())
        
        zip_buffer = io.BytesIO()
        with zipfile.ZipFile(zip_buffer, 'w', zipfile.ZIP_DEFLATED) as zip_file:
            for i, image in enumerate(images):
                img_byte_arr = io.BytesIO()
                image.save(img_byte_arr, format=format)
                img_byte_arr.seek(0)
                zip_file.writestr(f'page_{i+1}.{format.lower()}', img_byte_arr.getvalue())
        
        zip_buffer.seek(0)
        return zip_buffer
    except Exception as e:
        st.error(f"Error: {str(e)}")
        return None

def merge_pdfs(uploaded_files):
    """Merge multiple PDFs"""
    try:
        pdf_writer = PyPDF2.PdfWriter()
        
        for uploaded_file in uploaded_files:
            pdf_reader = PyPDF2.PdfReader(uploaded_file)
            for page in pdf_reader.pages:
                pdf_writer.add_page(page)
        
        output = io.BytesIO()
        pdf_writer.write(output)
        output.seek(0)
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
            # First part
            pdf_writer1 = PyPDF2.PdfWriter()
            for i in range(min(split_at, len(pdf_reader.pages))):
                pdf_writer1.add_page(pdf_reader.pages[i])
            
            output1 = io.BytesIO()
            pdf_writer1.write(output1)
            output1.seek(0)
            zip_file.writestr('part1.pdf', output1.getvalue())
            
            # Second part
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

def convert_image_format(uploaded_file, output_format):
    """Convert between image formats"""
    try:
        img = Image.open(uploaded_file)
        
        if output_format == 'JPEG' and img.mode in ('RGBA', 'P'):
            img = img.convert('RGB')
        
        output = io.BytesIO()
        img.save(output, format=output_format)
        output.seek(0)
        return output
    except Exception as e:
        st.error(f"Error: {str(e)}")
        return None

# Main conversion interface
st.markdown(f"### {conversion_type}")

# Conversion logic based on selected type
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

elif conversion_type == "Split PDF":
    uploaded_file = st.file_uploader("Upload PDF", type=['pdf'])
    if uploaded_file:
        pdf_reader = PyPDF2.PdfReader(uploaded_file)
        total_pages = len(pdf_reader.pages)
        st.info(f"PDF has {total_pages} pages")
        
        split_at = st.number_input("Split at page number", min_value=1, max_value=total_pages-1, value=1)
        
        if st.button("Split PDF"):
            with st.spinner("Splitting..."):
                result = split_pdf(uploaded_file, split_at)
                if result:
                    st.success("✅ PDF split successfully!")
                    st.download_button("📥 Download Split PDFs (ZIP)", result, f"{Path(uploaded_file.name).stem}_split.zip", "application/zip")

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

else:
    st.info(f"🚧 {conversion_type} is coming soon!")
    st.markdown("This conversion feature is under development.")

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
    st.markdown("- Word, Excel")
    st.markdown("- JPG, PNG Images")

with col3:
    st.markdown("**Other:**")
    st.markdown("- Merge/Split PDF")
    st.markdown("- Image Formats")
    st.markdown("- CSV ↔ Excel")



    # Core requirements
# pip install streamlit pillow pandas openpyxl python-docx python-pptx PyPDF2 pdf2image img2pdf reportlab poppler-utils



