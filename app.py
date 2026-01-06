import streamlit as st
import os
import tempfile
from markitdown import MarkItDown

# --- Configuration & Page Setup ---
st.set_page_config(
    page_title="Universal Document to Markdown Converter",
    page_icon="📝",
    layout="centered"
)

# Custom CSS for a cleaner look
st.markdown("""
    <style>
    .stTextArea textarea {
        font-family: 'Courier New', Courier, monospace;
        background-color: #f0f2f6;
    }
    .success-box {
        padding: 1rem;
        border-radius: 0.5rem;
        background-color: #d4edda;
        color: #155724;
        margin-bottom: 1rem;
    }
    </style>
""", unsafe_allow_html=True)

# --- The Engine Class ---
class DocumentConverter:
    def __init__(self):
        # Initialize MarkItDown
        self.md = MarkItDown()

    def process_file(self, uploaded_file):
        """
        Saves uploaded file temporarily, converts it, and cleans up.
        Returns: (success_boolean, content_or_error_message)
        """
        # Create a temp file to store the uploaded binary data
        # We preserve the suffix so MarkItDown knows how to handle it (e.g., .docx)
        suffix = os.path.splitext(uploaded_file.name)[1]
        
        with tempfile.NamedTemporaryFile(delete=False, suffix=suffix) as tmp_file:
            tmp_file.write(uploaded_file.getvalue())
            tmp_path = tmp_file.name

        try:
            # conversion_result is a Document object, we need .text_content
            # Note: MarkItDown handles requests internally, but for local files, 
            # we rely on standard file IO.
            result = self.md.convert(tmp_path)
            
            # If result is None or empty, raise an error
            if not result or not result.text_content:
                raise ValueError("Conversion yielded empty result.")
                
            return True, result.text_content

        except Exception as e:
            return False, str(e)
            
        finally:
            # Resilience: Ensure temp file is always deleted to prevent server clutter
            if os.path.exists(tmp_path):
                os.remove(tmp_path)

# --- Main Application Logic ---
def main():
    st.title("📝 Doc-to-MD Converter")
    st.markdown("Upload **PDF, Word, Excel, PowerPoint, or HTML** files to extract clean Markdown text.")

    # 1. Upload Area
    uploaded_files = st.file_uploader(
        "Drag and drop files here", 
        type=['pptx', 'docx', 'xlsx', 'pdf', 'html', 'htm'], 
        accept_multiple_files=True
    )

    if uploaded_files:
        converter = DocumentConverter()
        st.write("---")

        for uploaded_file in uploaded_files:
            with st.container():
                st.subheader(f"📄 Processing: {uploaded_file.name}")
                
                # Process the file
                with st.spinner(f"Reading {uploaded_file.name}..."):
                    success, content = converter.process_file(uploaded_file)

                # 2. Resilience & Error Handling
                if not success:
                    st.error(f"⚠️ Could not read [{uploaded_file.name}]. Please check the format. Error details: {content}")
                else:
                    # Success UI
                    base_name = os.path.splitext(uploaded_file.name)[0]
                    output_filename = f"{base_name}_converted"
                    
                    # 3. Instant Preview (Scrollable)
                    st.success("Conversion successful!")
                    st.text_area("Preview", value=content, height=300, key=f"preview_{uploaded_file.name}")

                    # 4. Download Options
                    col1, col2 = st.columns(2)
                    
                    # Option A: Download as Markdown (.md)
                    with col1:
                        st.download_button(
                            label="⬇️ Download Markdown (.md)",
                            data=content,
                            file_name=f"{output_filename}.md",
                            mime="text/markdown",
                            key=f"dl_md_{uploaded_file.name}"
                        )
                    
                    # Option B: Download as Text (.txt)
                    with col2:
                        st.download_button(
                            label="⬇️ Download Text (.txt)",
                            data=content,
                            file_name=f"{output_filename}.txt",
                            mime="text/plain",
                            key=f"dl_txt_{uploaded_file.name}"
                        )
                
                st.write("---")

if __name__ == "__main__":
    main()
