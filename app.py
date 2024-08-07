import streamlit as st
import pytesseract
import pandas as pd
from PIL import Image
from sqlalchemy import create_engine, Column, Integer, String, Text
from sqlalchemy.ext.declarative import declarative_base
from sqlalchemy.orm import sessionmaker
from sqlalchemy import inspect
from docx import Document
from docx2pdf import convert
import pdf2image
import os
import tempfile

# SQLAlchemy setup
db_url = 'sqlite:///extracted_text_new.db'
engine = create_engine(db_url)
Base = declarative_base()


class ExtractedTextNew(Base):
    __tablename__ = 'extracted_text_new'
    id = Column(Integer, primary_key=True, autoincrement=True)
    file_name = Column(String)
    content = Column(Text)


# Drop the existing table if it exists
inspector = inspect(engine)
if inspector.has_table(ExtractedTextNew.__tablename__):
    Base.metadata.drop_all(bind=engine, tables=[ExtractedTextNew.__table__])

Base.metadata.create_all(engine)
Session = sessionmaker(bind=engine)
session = Session()


def extract_text_from_image_with_tesseract(image):
    custom_config = r'--oem 3 --psm 6'
    return pytesseract.image_to_string(image, config=custom_config)


def process_pdf(uploaded_file):
    st.write("Processing PDF file...")
    with open(uploaded_file.name, "wb") as f:
        f.write(uploaded_file.getbuffer())

    try:
        images = pdf2image.convert_from_path(uploaded_file.name, dpi=300)
        content = "\n".join(extract_text_from_image_with_tesseract(image) for image in images)
        page_text = ExtractedTextNew(file_name=uploaded_file.name, content=content)
        session.add(page_text)
        session.commit()
        st.success(f"Successfully processed PDF: {uploaded_file.name}")
    except Exception as e:
        st.error(f"Error processing PDF file: {e}")


def process_excel(uploaded_file):
    st.write("Processing Excel file...")
    df = pd.read_excel(uploaded_file, sheet_name=None)
    for sheet_name, data in df.items():
        st.write(f"Processing sheet: {sheet_name}")
        content = data.to_string(index=False)
        excel_text = ExtractedTextNew(file_name=uploaded_file.name, content=content)
        session.add(excel_text)
    session.commit()


def process_word_with_ocr(uploaded_file):
    st.write("Processing Word file with OCR...")
    temp_dir = tempfile.mkdtemp()
    temp_pdf_path = os.path.join(temp_dir, "temp.pdf")

    try:
        with open(uploaded_file.name, "wb") as f:
            f.write(uploaded_file.getbuffer())

        # Convert Word document to PDF
        convert(uploaded_file.name, temp_pdf_path)

        # Convert PDF pages to images
        images = pdf2image.convert_from_path(temp_pdf_path, dpi=300)
        content = "\n".join(extract_text_from_image_with_tesseract(image) for image in images)

        word_text = ExtractedTextNew(file_name=uploaded_file.name, content=content)
        session.add(word_text)
        session.commit()

        st.success(f"Successfully processed Word file with OCR: {uploaded_file.name}")
    except Exception as e:
        st.error(f"Error processing Word file with OCR: {e}")
    finally:
        os.remove(temp_pdf_path)
        os.rmdir(temp_dir)


def download_text_file():
    results = session.query(ExtractedTextNew).all()
    combined_content = "\n\n".join([f"File: {result.file_name}\n\n{result.content}" for result in results])
    st.download_button(
        label="Download Extracted Text",
        data=combined_content,
        file_name="extracted_text.txt",
        mime="text/plain"
    )


st.title("File Upload and Processing")
uploaded_files = st.file_uploader("Upload PDF, Excel, or Word files", type=["pdf", "xlsx", "docx"],
                                  accept_multiple_files=True)
process_button = st.button("Process Files")

if process_button:
    if uploaded_files:
        for uploaded_file in uploaded_files:
            file_type = uploaded_file.name.split('.')[-1]
            if file_type == 'pdf':
                process_pdf(uploaded_file)
            elif file_type == 'xlsx':
                process_excel(uploaded_file)
            elif file_type == 'docx':
                process_word_with_ocr(uploaded_file)
        st.success("Files processed successfully.")
        download_text_file()
    else:
        st.warning("Please upload files to process.")
