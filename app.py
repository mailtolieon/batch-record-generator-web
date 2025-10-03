import streamlit as st
import pandas as pd
import os
import tempfile
import base64
from datetime import datetime
import zipfile
import io
from pathlib import Path
import time

from document_generator import WebDocumentGenerator
from config import WebConfig

# Page configuration
st.set_page_config(
    page_title="Batch Record Generator",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Custom CSS
def load_css():
    st.markdown("""
    <style>
    .main-header {
        font-size: 2.5rem;
        color: #1f77b4;
        text-align: center;
        margin-bottom: 2rem;
    }
    .success-box {
        padding: 1rem;
        border-radius: 0.5rem;
        background-color: #d4edda;
        border: 1px solid #c3e6cb;
        color: #155724;
    }
    .error-box {
        padding: 1rem;
        border-radius: 0.5rem;
        background-color: #f8d7da;
        border: 1px solid #f5c6cb;
        color: #721c24;
    }
    .info-box {
        padding: 1rem;
        border-radius: 0.5rem;
        background-color: #d1ecf1;
        border: 1px solid #bee5eb;
        color: #0c5460;
    }
    .product-card {
        padding: 0.5rem;
        margin: 0.25rem 0;
        border-radius: 0.25rem;
        background-color: #f8f9fa;
        border: 1px solid #dee2e6;
    }
    </style>
    """, unsafe_allow_html=True)

def init_session_state():
    """Initialize session state variables"""
    if 'doc_generator' not in st.session_state:
        st.session_state.doc_generator = None
    if 'products' not in st.session_state:
        st.session_state.products = []
    if 'selected_products' not in st.session_state:
        st.session_state.selected_products = []
    if 'generation_history' not in st.session_state:
        st.session_state.generation_history = []
    if 'data_loaded' not in st.session_state:
        st.session_state.data_loaded = False

def setup_data_directories():
    """Create necessary directories"""
    directories = ['data', 'generated', 'temp']
    for directory in directories:
        Path(directory).mkdir(exist_ok=True)

def get_file_download_link(file_path, filename):
    """Create a download link for a file"""
    with open(file_path, "rb") as f:
        data = f.read()
    b64 = base64.b64encode(data).decode()
    href = f'<a href="data:application/octet-stream;base64,{b64}" download="{filename}">Download {filename}</a>'
    return href

def create_zip_download(files):
    """Create a ZIP file for multiple downloads"""
    zip_buffer = io.BytesIO()
    with zipfile.ZipFile(zip_buffer, "w") as zip_file:
        for file_path, filename in files:
            zip_file.write(file_path, filename)
    zip_buffer.seek(0)
    return zip_buffer

def main():
    load_css()
    init_session_state()
    setup_data_directories()
    
    # Header
    st.markdown('<h1 class="main-header">📊 Batch Record Generator</h1>', unsafe_allow_html=True)
    
    # Sidebar
    with st.sidebar:
        st.header("⚙️ Configuration")
        
        # File uploaders
        st.subheader("Data Files")
        excel_file = st.file_uploader("Upload Excel File", type=['xlsx', 'xls'])
        template_file = st.file_uploader("Upload Word Template", type=['docx'])
        
        # Configuration options
        st.subheader("Settings")
        sheet_name = st.text_input("Sheet Name", value="5_Arc_List")
        auto_download = st.checkbox("Auto-download generated files", value=True)
        
        # Initialize button
        if st.button("🚀 Initialize Application", use_container_width=True):
            if excel_file and template_file:
                with st.spinner("Initializing application..."):
                    try:
                        # Save uploaded files
                        excel_path = f"data/{excel_file.name}"
                        template_path = f"data/{template_file.name}"
                        
                        with open(excel_path, "wb") as f:
                            f.write(excel_file.getvalue())
                        with open(template_path, "wb") as f:
                            f.write(template_file.getvalue())
                        
                        # Initialize document generator
                        config = WebConfig(
                            excel_file=excel_path,
                            template_file=template_path,
                            sheet_name=sheet_name
                        )
                        
                        st.session_state.doc_generator = WebDocumentGenerator(config)
                        st.session_state.data_loaded = True
                        
                        # Load products
                        st.session_state.products = st.session_state.doc_generator.get_unique_products()
                        
                        st.success("✅ Application initialized successfully!")
                        
                    except Exception as e:
                        st.error(f"❌ Error initializing application: {str(e)}")
            else:
                st.warning("⚠️ Please upload both Excel and Template files")
    
    # Main content area
    if not st.session_state.data_loaded:
        st.info("👈 Please upload your Excel file and Word template in the sidebar to get started.")
        
        # Quick start guide
        with st.expander("📖 Quick Start Guide"):
            st.markdown("""
            ### How to use this application:
            
            1. **Upload Files**: Use the sidebar to upload your Excel data file and Word template
            2. **Initialize**: Click the 'Initialize Application' button
            3. **Select Products**: Choose products from the list
            4. **Generate**: Create single or multiple documents
            5. **Download**: Get your generated Word documents
            
            ### Template Requirements:
            - Word document must contain `{{PRODUCT_NAME}}` placeholder
            - Should include a table for batch data
            - Supports headers and footers
            
            ### Excel Requirements:
            - Should contain product and batch information
            - Column names are automatically detected
            """)
        
        return
    
    # Main application interface
    tab1, tab2, tab3, tab4 = st.tabs(["📋 Product Selection", "🚀 Document Generation", "📊 Batch Preview", "📈 History & Analytics"])
    
    with tab1:
        st.header("Product Selection")
        
        # Search and filter
        col1, col2 = st.columns([2, 1])
        
        with col1:
            search_query = st.text_input("🔍 Search Products", placeholder="Type to search...")
        
        with col2:
            filter_option = st.selectbox(
                "Filter by",
                ["All Products", "With Batches", "No Batches"]
            )
        
        # Apply filters
        filtered_products = st.session_state.products
        if search_query:
            filtered_products = [p for p in filtered_products if search_query.lower() in p.lower()]
        
        # Product selection
        st.subheader(f"Available Products ({len(filtered_products)})")
        
        if filtered_products:
            # Select all checkbox
            col1, col2, col3 = st.columns([1, 1, 2])
            with col1:
                if st.button("Select All", use_container_width=True):
                    st.session_state.selected_products = filtered_products.copy()
            with col2:
                if st.button("Clear Selection", use_container_width=True):
                    st.session_state.selected_products = []
            
            # Product selection checkboxes
            selected = []
            for product in filtered_products:
                is_selected = product in st.session_state.selected_products
                if st.checkbox(product, value=is_selected, key=f"prod_{product}"):
                    selected.append(product)
            
            st.session_state.selected_products = selected
            
            st.info(f"📦 {len(selected)} product(s) selected")
            
        else:
            st.warning("No products found matching your criteria.")
    
    with tab2:
        st.header("Document Generation")
        
        if not st.session_state.selected_products:
            st.warning("Please select at least one product from the Product Selection tab.")
        else:
            col1, col2 = st.columns(2)
            
            with col1:
                st.subheader("Single Document")
                selected_product = st.selectbox(
                    "Choose a product for single generation",
                    st.session_state.selected_products
                )
                
                if st.button("🔄 Generate Single Document", use_container_width=True):
                    with st.spinner(f"Generating document for {selected_product}..."):
                        try:
                            result = st.session_state.doc_generator.generate_single_document(selected_product)
                            
                            if result['success']:
                                st.success(f"✅ Document generated: {result['filename']}")
                                
                                # Add to history
                                st.session_state.generation_history.append({
                                    'timestamp': datetime.now(),
                                    'product': selected_product,
                                    'batches': result['batch_count'],
                                    'filename': result['filename'],
                                    'type': 'single'
                                })
                                
                                # Download link
                                if auto_download:
                                    download_link = get_file_download_link(result['filepath'], result['filename'])
                                    st.markdown(download_link, unsafe_allow_html=True)
                                
                                # Show batch info
                                st.info(f"📊 Contains {result['batch_count']} batches")
                                
                            else:
                                st.error(f"❌ Failed to generate document: {result['error']}")
                                
                        except Exception as e:
                            st.error(f"❌ Error generating document: {str(e)}")
            
            with col2:
                st.subheader("Bulk Generation")
                st.info(f"Will generate documents for {len(st.session_state.selected_products)} selected products")
                
                if st.button("🚀 Generate Bulk Documents", use_container_width=True):
                    progress_bar = st.progress(0)
                    status_text = st.empty()
                    
                    results = []
                    total_products = len(st.session_state.selected_products)
                    
                    for i, product in enumerate(st.session_state.selected_products):
                        status_text.text(f"Processing {product} ({i+1}/{total_products})")
                        progress_bar.progress((i + 1) / total_products)
                        
                        try:
                            result = st.session_state.doc_generator.generate_single_document(product)
                            results.append(result)
                            
                            # Add to history
                            if result['success']:
                                st.session_state.generation_history.append({
                                    'timestamp': datetime.now(),
                                    'product': product,
                                    'batches': result['batch_count'],
                                    'filename': result['filename'],
                                    'type': 'bulk'
                                })
                            
                            time.sleep(0.1)  # Small delay for UI update
                            
                        except Exception as e:
                            results.append({
                                'success': False,
                                'product': product,
                                'error': str(e)
                            })
                    
                    # Show results
                    successful = [r for r in results if r['success']]
                    failed = [r for r in results if not r['success']]
                    
                    if successful:
                        st.success(f"✅ Successfully generated {len(successful)} documents")
                        
                        # Create ZIP download for all successful files
                        if len(successful) > 1 and auto_download:
                            files_to_zip = [(r['filepath'], r['filename']) for r in successful]
                            zip_buffer = create_zip_download(files_to_zip)
                            
                            st.download_button(
                                label="📦 Download All as ZIP",
                                data=zip_buffer,
                                file_name=f"batch_documents_{datetime.now().strftime('%Y%m%d_%H%M%S')}.zip",
                                mime="application/zip",
                                use_container_width=True
                            )
                    
                    if failed:
                        st.error(f"❌ Failed to generate {len(failed)} documents")
                        with st.expander("Show failed generations"):
                            for fail in failed:
                                st.write(f"- {fail['product']}: {fail['error']}")
    
    with tab3:
        st.header("Batch Preview")
        
        if st.session_state.selected_products:
            preview_product = st.selectbox(
                "Select product to preview",
                st.session_state.selected_products
            )
            
            if st.button("🔍 Preview Batch Data"):
                with st.spinner("Loading batch data..."):
                    try:
                        preview_data = st.session_state.doc_generator.preview_product_data(preview_product)
                        
                        if preview_data:
                            st.success(f"📊 Found {len(preview_data)} batches for {preview_product}")
                            
                            # Display as dataframe
                            df = pd.DataFrame(preview_data)
                            st.dataframe(df, use_container_width=True)
                            
                            # Basic statistics
                            col1, col2, col3, col4 = st.columns(4)
                            with col1:
                                st.metric("Total Batches", len(preview_data))
                            with col2:
                                valid_dates = len([b for b in preview_data if b.get('mfg_date')])
                                st.metric("With MFG Dates", valid_dates)
                            with col3:
                                valid_yield = len([b for b in preview_data if b.get('total_batch_yield')])
                                st.metric("With Yield Data", valid_yield)
                            with col4:
                                sent_to_doc = len([b for b in preview_data if b.get('sent_to_document_room')])
                                st.metric("Sent to Doc Room", sent_to_doc)
                            
                        else:
                            st.warning("No batch data found for this product.")
                            
                    except Exception as e:
                        st.error(f"Error loading preview data: {str(e)}")
        else:
            st.info("Please select products in the Product Selection tab to preview batch data.")
    
    with tab4:
        st.header("Generation History & Analytics")
        
        if st.session_state.generation_history:
            # Convert history to dataframe
            history_df = pd.DataFrame(st.session_state.generation_history)
            history_df['timestamp'] = pd.to_datetime(history_df['timestamp'])
            
            # Display history
            st.subheader("Recent Generations")
            st.dataframe(
                history_df.sort_values('timestamp', ascending=False).head(10),
                use_container_width=True
            )
            
            # Analytics
            col1, col2, col3 = st.columns(3)
            
            with col1:
                total_generated = len(history_df)
                st.metric("Total Documents", total_generated)
            
            with col2:
                unique_products = history_df['product'].nunique()
                st.metric("Unique Products", unique_products)
            
            with col3:
                total_batches = history_df['batches'].sum()
                st.metric("Total Batches", total_batches)
            
            # Recent activity chart
            st.subheader("Generation Activity")
            daily_counts = history_df.groupby(history_df['timestamp'].dt.date).size()
            st.bar_chart(daily_counts)
            
        else:
            st.info("No generation history yet. Generate some documents to see analytics here.")

if __name__ == "__main__":
    main()