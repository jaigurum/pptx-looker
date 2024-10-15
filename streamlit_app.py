import fitz  # PyMuPDF
from pptx import Presentation
from pptx.util import Inches
import streamlit as st
import tempfile
import os

# Function to convert each PDF page to an image and insert into PPTX
def convert_pdf_to_pptx_from_pages(pdf_file, pdf_filename):
    # Open the uploaded PDF using PyMuPDF
    pdf_document = fitz.open(stream=pdf_file.read(), filetype="pdf")
    
    # Create a new PowerPoint presentation
    presentation = Presentation()
    
    # Iterate over each page in the PDF
    for page_num in range(pdf_document.page_count):
        page = pdf_document.load_page(page_num)  # Load each page
        
        # Render the page to an image (resolution 150 dpi for decent quality)
        pix = page.get_pixmap(dpi=150)
        
        # Save the image as a temporary file
        with tempfile.NamedTemporaryFile(delete=False, suffix=".png") as img_tmp:
            img_tmp.write(pix.tobytes())
            img_tmp_path = img_tmp.name
            
            # Add a blank slide to the presentation
            slide_layout = presentation.slide_layouts[6]  # Blank layout
            slide = presentation.slides.add_slide(slide_layout)
            
            # Insert the image into the slide (size: full slide)
            slide.shapes.add_picture(img_tmp_path, Inches(0), Inches(0), width=Inches(10), height=Inches(7.5))
    
    # Save the pptx to a temporary file with the same name as the PDF but with .pptx extension
    pptx_filename = os.path.splitext(pdf_filename)[0] + ".pptx"
    with tempfile.NamedTemporaryFile(delete=False, suffix=".pptx") as tmp_file:
        pptx_file = tmp_file.name
        presentation.save(pptx_file)
    
    return pptx_file, pptx_filename

# Streamlit app UI
def main():
    st.title("VML PDF to PPTX Converter")

    # Step 1: Upload PDF
    uploaded_pdf = st.file_uploader("Upload your lookerstudio report as a PDF file", type="pdf")

    if uploaded_pdf is not None:
        st.success("PDF uploaded successfully!")
        
        # Extract the filename from the uploaded file
        pdf_filename = uploaded_pdf.name

        # Step 2: Convert to PPTX
        if st.button("Convert to PPTX"):
            # Pass the uploaded PDF to the conversion function
            with st.spinner("Converting..."):
                pptx_path, pptx_filename = convert_pdf_to_pptx_from_pages(uploaded_pdf, pdf_filename)
                st.success(f"Conversion successful! each page in pdf is in a separate slide. Download {pptx_filename}")

                # Step 3: Provide download link
                with open(pptx_path, "rb") as pptx_file:
                    st.download_button(
                        label="Download PPTX",
                        data=pptx_file,
                        file_name=pptx_filename,
                        mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
                    )

if __name__ == "__main__":
    main()
