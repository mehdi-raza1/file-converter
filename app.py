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
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Image as RLImage, Table, TableStyle, PageBreak
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import inch
from reportlab.lib import colors
from reportlab.lib.enums import TA_CENTER, TA_LEFT, TA_RIGHT, TA_JUSTIFY
from reportlab.lib.utils import ImageReader

# Set page configuration
st.set_page_config(
    page_title="Universal File Converter Proooo",
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
        
        pdf_doc.build(story)
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
    """Convert PowerPoint to PDF"""
    try:
        prs = Presentation(uploaded_file)
        
        output = io.BytesIO()
        pdf_doc = SimpleDocTemplate(output, pagesize=letter)
        styles = getSampleStyleSheet()
        story = []
        
        title_style = ParagraphStyle(
            'SlideTitle',
            parent=styles['Heading1'],
            fontSize=24,
            textColor=colors.HexColor('#4472C4'),
            spaceAfter=20,
            alignment=TA_CENTER
        )
        
        for i, slide in enumerate(prs.slides):
            story.append(Paragraph(f"<b>Slide {i+1}</b>", title_style))
            story.append(Spacer(1, 12))
            
            for shape in slide.shapes:
                if hasattr(shape, "text") and shape.text:
                    text = shape.text.strip()
                    if text:
                        story.append(Paragraph(text, styles['Normal']))
                        story.append(Spacer(1, 8))
                
                if shape.shape_type == 13:  # Picture
                    try:
                        image = shape.image
                        image_bytes = image.blob
                        
                        img = ImageReader(io.BytesIO(image_bytes))
                        img_width, img_height = img.getSize()
                        
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
                
                if hasattr(shape, 'has_table') and shape.has_table:
                    table_data = []
                    for row in shape.table.rows:
                        row_data = [cell.text for cell in row.cells]
                        table_data.append(row_data)
                    
                    if table_data:
                        t = Table(table_data)
                        t.setStyle(TableStyle([
                            ('BACKGROUND', (0, 0), (-1, 0), colors.grey),
                            ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
                            ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
                            ('GRID', (0, 0), (-1, -1), 1, colors.black),
                        ]))
                        story.append(t)
                        story.append(Spacer(1, 12))
            
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