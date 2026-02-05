import streamlit as st
from openai import OpenAI
import os
from datetime import datetime
import json
import io

# Import document parsing libraries
try:
    import PyPDF2
except ImportError:
    st.warning("PyPDF2 not installed. PDF support unavailable.")

try:
    from docx import Document
except ImportError:
    st.warning("python-docx not installed. DOCX support unavailable.")

try:
    import openpyxl
    import pandas as pd
except ImportError:
    st.warning("openpyxl/pandas not installed. Excel support unavailable.")

# Page configuration
st.set_page_config(
    page_title="Document Chat",
    page_icon="📄",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Custom CSS for better UI
st.markdown("""
    <style>
    .main { padding: 2rem 1rem; }
    .stChatMessage { margin: 1rem 0; }
    [data-testid="stSidebar"] { background-color: #f0f2f6; }
    .document-info { background: #e3f2fd; padding: 1rem; border-radius: 0.5rem; margin: 1rem 0; }
    .chat-container { max-width: 900px; margin: 0 auto; }
    .file-type-badge { display: inline-block; padding: 0.25rem 0.5rem; background: #4CAF50; color: white; border-radius: 0.25rem; font-size: 0.8rem; margin-left: 0.5rem; }
    </style>
""", unsafe_allow_html=True)

# Initialize session state
if "messages" not in st.session_state:
    st.session_state.messages = []
if "document_content" not in st.session_state:
    st.session_state.document_content = None
if "document_name" not in st.session_state:
    st.session_state.document_name = None
if "document_type" not in st.session_state:
    st.session_state.document_type = None
if "saved_contexts" not in st.session_state:
    st.session_state.saved_contexts = {}
if "client" not in st.session_state:
    st.session_state.client = None

@st.cache_resource
def initialize_openai_client(api_key):
    """Initialize and cache OpenAI client"""
    return OpenAI(api_key=api_key)

@st.cache_data
def process_document(file_content: str, file_name: str, file_type: str):
    """Cache processed document content"""
    return {
        "content": file_content,
        "name": file_name,
        "type": file_type,
        "processed_at": datetime.now().isoformat()
    }

def extract_pdf_content(file_object) -> str:
    """Extract text content from PDF file"""
    try:
        pdf_reader = PyPDF2.PdfReader(file_object)
        text = ""
        for page_num in range(len(pdf_reader.pages)):
            page = pdf_reader.pages[page_num]
            text += f"\n--- Page {page_num + 1} ---\n"
            text += page.extract_text()
        return text
    except Exception as e:
        raise Exception(f"Error extracting PDF: {str(e)}")

def extract_docx_content(file_object) -> str:
    """Extract text content from DOCX file"""
    try:
        doc = Document(file_object)
        text = ""
        for para in doc.paragraphs:
            if para.text.strip():
                text += para.text + "\n"
        
        # Extract tables if any
        for table in doc.tables:
            text += "\n[TABLE]\n"
            for row in table.rows:
                row_text = " | ".join([cell.text for cell in row.cells])
                text += row_text + "\n"
        
        return text
    except Exception as e:
        raise Exception(f"Error extracting DOCX: {str(e)}")

def extract_excel_content(file_object) -> str:
    """Extract content from Excel file (XLS/XLSX)"""
    try:
        excel_file = pd.ExcelFile(file_object)
        text = ""
        
        for sheet_name in excel_file.sheet_names:
            df = pd.read_excel(file_object, sheet_name=sheet_name)
            text += f"\n--- Sheet: {sheet_name} ---\n"
            text += df.to_string()
            text += "\n"
        
        return text
    except Exception as e:
        raise Exception(f"Error extracting Excel: {str(e)}")

def extract_text_content(file_object) -> str:
    """Extract text from TXT or MD files"""
    try:
        return file_object.read().decode('utf-8')
    except Exception as e:
        raise Exception(f"Error reading text file: {str(e)}")

def save_context(context_name: str):
    """Save current conversation context"""
    if st.session_state.messages:
        st.session_state.saved_contexts[context_name] = {
            "messages": st.session_state.messages.copy(),
            "document": st.session_state.document_name,
            "document_type": st.session_state.document_type,
            "saved_at": datetime.now().isoformat()
        }
        st.success(f"Context '{context_name}' saved successfully!")
    else:
        st.warning("No conversation to save.")

def load_context(context_name: str):
    """Load saved conversation context"""
    if context_name in st.session_state.saved_contexts:
        context = st.session_state.saved_contexts[context_name]
        st.session_state.messages = context["messages"].copy()
        st.success(f"Context '{context_name}' loaded successfully!")
        st.rerun()
    else:
        st.error("Context not found.")

def clear_chat():
    """Clear current conversation"""
    st.session_state.messages = []
    st.rerun()

# Sidebar configuration
with st.sidebar:
    st.title("⚙️ Configuration")
    
    api_key = st.text_input("OpenAI API Key", type="password", help="Enter your OpenAI API key")
    
    if api_key:
        st.session_state.client = initialize_openai_client(api_key)
    
    st.divider()
    
    # Document upload section
    st.subheader("📁 Document Upload")
    st.caption("Supported formats: TXT, MD, PDF, DOCX, XLS, XLSX")
    
    uploaded_file = st.file_uploader(
        "Upload a document",
        type=["txt", "md", "pdf", "docx", "doc", "xls", "xlsx"],
        accept_multiple_files=False
    )
    
    if uploaded_file:
        try:
            file_extension = uploaded_file.name.split('.')[-1].lower()
            file_content = None
            
            # Reset file pointer to beginning
            uploaded_file.seek(0)
            
            if file_extension == "pdf":
                file_content = extract_pdf_content(uploaded_file)
                doc_type = "PDF"
            elif file_extension in ["docx", "doc"]:
                file_content = extract_docx_content(uploaded_file)
                doc_type = "DOCX"
            elif file_extension in ["xls", "xlsx"]:
                file_content = extract_excel_content(uploaded_file)
                doc_type = "Excel"
            elif file_extension in ["txt", "md"]:
                file_content = extract_text_content(uploaded_file)
                doc_type = "Text"
            
            if file_content:
                processed = process_document(file_content, uploaded_file.name, doc_type)
                st.session_state.document_content = processed["content"]
                st.session_state.document_name = processed["name"]
                st.session_state.document_type = processed["type"]
                st.success(f"✅ Document loaded: {uploaded_file.name}")
        
        except Exception as e:
            st.error(f"Error loading document: {e}")
    
    if st.session_state.document_name:
        doc_type_badge = f'<span class="file-type-badge">{st.session_state.document_type}</span>'
        st.markdown(f"""
        <div class="document-info">
        <strong>Current Document:</strong><br>
        📄 {st.session_state.document_name} {doc_type_badge}<br>
        <small>Size: {len(st.session_state.document_content) / 1024:.2f} KB</small>
        </div>
        """, unsafe_allow_html=True)
    
    st.divider()
    
    # Context management
    st.subheader("💾 Context Management")
    
    col1, col2 = st.columns(2)
    with col1:
        if st.button("💾 Save Context", use_container_width=True):
            st.session_state.show_save_dialog = True
    
    with col2:
        if st.button("🗑️ Clear Chat", use_container_width=True):
            clear_chat()
    
    if "show_save_dialog" in st.session_state and st.session_state.show_save_dialog:
        context_name = st.text_input("Context name:", key="context_name_input")
        if st.button("Save", key="save_btn"):
            if context_name:
                save_context(context_name)
                st.session_state.show_save_dialog = False
                st.rerun()
    
    if st.session_state.saved_contexts:
        st.subheader("📋 Saved Contexts")
        for context_name in st.session_state.saved_contexts.keys():
            col1, col2 = st.columns([3, 1])
            with col1:
                if st.button(f"📂 {context_name}", use_container_width=True, key=f"load_{context_name}"):
                    load_context(context_name)
            with col2:
                if st.button("×", key=f"delete_{context_name}"):
                    del st.session_state.saved_contexts[context_name]
                    st.rerun()

# Main chat interface
st.title("📄 Document Chat Assistant")

if not st.session_state.client:
    st.warning("⚠️ Please enter your OpenAI API key in the sidebar to proceed.")
else:
    if not st.session_state.document_content:
        st.info("📤 Upload a document in the sidebar to start chatting.")
    else:
        # Display chat messages
        for message in st.session_state.messages:
            with st.chat_message(message["role"]):
                st.markdown(message["content"])
        
        # Chat input
        if prompt := st.chat_input("Ask something about your document..."):
            if not st.session_state.document_content:
                st.error("Please upload a document first.")
            else:
                # Add user message
                st.session_state.messages.append({"role": "user", "content": prompt})
                
                with st.chat_message("user"):
                    st.markdown(prompt)
                
                # Generate response with streaming
                with st.chat_message("assistant"):
                    try:
                        system_prompt = f"""You are a helpful assistant. Answer questions based ONLY on the following document content.
                        
DOCUMENT CONTENT:
---
{st.session_state.document_content}
---

Be concise and helpful. If the answer is not in the document, say so."""
                        
                        with st.spinner("Thinking..."):
                            stream = st.session_state.client.chat.completions.create(
                                model="gpt-4o",
                                messages=[
                                    {"role": "system", "content": system_prompt},
                                    *st.session_state.messages
                                ],
                                stream=True,
                                temperature=0.7,
                                max_tokens=1000
                            )
                            
                            response_content = st.write_stream(stream)
                        
                        # Add assistant message
                        st.session_state.messages.append({"role": "assistant", "content": response_content})
                    
                    except Exception as e:
                        st.error(f"Error generating response: {e}")