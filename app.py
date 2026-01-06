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

# Custom CSS
st.markdown("""
    <style>
    .stTextArea textarea {
        font-family: 'Courier New', Courier, monospace;
        background-color: #f0f2f6;
    }
    .metric-card {
        background-color: #f8f9fa;
        border: 1px solid #e9ecef;
        padding: 1rem;
        border-radius: 0.5rem;
        margin-bottom: 0.5rem;
        text-align: center;
    }
    .highlight-red { color: #dc3545; font-weight: bold; }
    .highlight-green { color: #28a745; font-weight: bold; }
    </style>
""", unsafe_allow_html=True)

# --- Helper: Format File Size ---
def format_size(size_in_bytes):
    for unit in ['B', 'KB', 'MB', 'GB']:
        if size_in_bytes < 1024.0:
            return f"{size_in_bytes:.2f} {unit}"
        size_in_bytes /= 1024.0
    return f"{size_in_bytes:.2f} TB"

# --- The Engine Class ---
class DocumentConverter:
    def __init__(self):
        self.md = MarkItDown()

    def process_file(self, uploaded_file):
        """
        Saves uploaded file temporarily, converts it, and cleans up.
        Returns: (success_boolean, content_or_error_message)
        """
        suffix = os.path.splitext(uploaded_file.name)[1]
        
        with tempfile.NamedTemporaryFile(delete=False, suffix=suffix) as tmp_file:
            tmp_file.write(uploaded_file.getvalue())
            tmp_path = tmp_file.name

        try:
            result = self.md.convert(tmp_path)
            if not result or not result.text_content:
                raise ValueError("Conversion yielded empty result.")
            return True, result.text_content

        except Exception as e:
            return False, str(e)
            
        finally:
            if os.path.exists(tmp_path):
                os.remove(tmp_path)

# --- Main Application Logic ---
def main():
    st.title("📝 Doc-to-MD Converter")

    # Initialize session state to store stats across tabs
    if 'conversion_stats' not in st.session_state:
        st.session_state.conversion_stats = []

    # Create Tabs
    tab1, tab2 = st.tabs(["🚀 Converter", "📊 File Size Comparison"])

    # --- TAB 1: The Converter ---
    with tab1:
        st.markdown("Upload **PDF, Word, Excel, PowerPoint, or HTML** files.")
        
        uploaded_files = st.file_uploader(
            "Drag and drop files here", 
            type=['pptx', 'docx', 'xlsx', 'pdf', 'html', 'htm'], 
            accept_multiple_files=True
        )

        if uploaded_files:
            converter = DocumentConverter()
            
            # Reset stats on new upload if you want fresh stats every time (optional)
            # st.session_state.conversion_stats = [] 

            for uploaded_file in uploaded_files:
                # Check if we already processed this specific file to avoid re-processing on refresh
                # (Simple check by name, usually sufficient for this demo)
                already_processed = any(s['name'] == uploaded_file.name for s in st.session_state.conversion_stats)
                
                if not already_processed:
                    with st.spinner(f"Reading {uploaded_file.name}..."):
                        success, content = converter.process_file(uploaded_file)
                        
                        if success:
                            # Calculate Stats
                            original_size = uploaded_file.size
                            converted_size = len(content.encode('utf-8'))
                            
                            # Add to session state
                            st.session_state.conversion_stats.append({
                                'name': uploaded_file.name,
                                'original_size': original_size,
                                'converted_size': converted_size,
                                'content': content
                            })
                        else:
                            st.error(f"⚠️ Could not read [{uploaded_file.name}]. Error: {content}")

            # Display Results for all successful items in session state
            if st.session_state.conversion_stats:
                st.write("---")
                for item in st.session_state.conversion_stats:
                    with st.expander(f"✅ Result: {item['name']}", expanded=True):
                        st.text_area("Preview", value=item['content'], height=200, key=f"preview_{item['name']}")
                        
                        base_name = os.path.splitext(item['name'])[0]
                        col1, col2 = st.columns(2)
                        with col1:
                            st.download_button("⬇️ Markdown (.md)", item['content'], f"{base_name}.md", key=f"md_{item['name']}")
                        with col2:
                            st.download_button("⬇️ Text (.txt)", item['content'], f"{base_name}.txt", key=f"txt_{item['name']}")

    # --- TAB 2: File Size Comparison ---
    with tab2:
        if not st.session_state.conversion_stats:
            st.info("Upload and convert files in the 'Converter' tab to see the comparison here.")
        else:
            st.subheader("Compression Analysis")
            
            for item in st.session_state.conversion_stats:
                orig_fmt = format_size(item['original_size'])
                new_fmt = format_size(item['converted_size'])
                
                # Calculate percentage reduction
                if item['original_size'] > 0:
                    reduction = ((item['original_size'] - item['converted_size']) / item['original_size']) * 100
                    reduction_str = f"{reduction:.1f}% smaller"
                    color_class = "highlight-green" if reduction > 0 else "highlight-red"
                else:
                    reduction_str = "N/A"
                    color_class = ""

                # Display Logic
                st.markdown(f"**📄 File:** `{item['name']}`")
                
                # Create a 3-column layout for the table row
                c1, c2, c3 = st.columns(3)
                c1.metric("Original Size", orig_fmt)
                c2.metric("Text Size", new_fmt)
                c3.markdown(f"<div style='padding-top: 15px;' class='{color_class}'>{reduction_str}</div>", unsafe_allow_html=True)
                
                st.divider()

if __name__ == "__main__":
    main()
