"""
JSON to Excel Converter - Production Ready Tool
A professional Streamlit application for converting JSON/NDJSON records to Excel files.

Features:
- Supports JSON Lines (NDJSON), JSON arrays, and .records files
- Automatic nested JSON flattening
- Multiple sheet support for complex data structures
- Data preview and statistics
- Professional UI/UX with dark theme support

Author: Data Engineering Team
Version: 1.0.0
"""

import io
import json
import re
from datetime import datetime
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple, Union
from collections import OrderedDict

import streamlit as st
import pandas as pd

# Page configuration
st.set_page_config(
    page_title="JSON to Excel Converter",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Custom CSS for professional look
st.markdown("""
<style>
    /* Main container styling */
    .main {
        padding: 1rem 2rem;
    }
    
    /* Header styling */
    .main-header {
        background: linear-gradient(135deg, #1e3a5f 0%, #2d5a87 100%);
        padding: 2rem;
        border-radius: 12px;
        margin-bottom: 2rem;
        color: white;
        text-align: center;
    }
    
    .main-header h1 {
        margin: 0;
        font-size: 2.5rem;
        font-weight: 700;
    }
    
    .main-header p {
        margin: 0.5rem 0 0 0;
        opacity: 0.9;
        font-size: 1.1rem;
    }
    
    /* Card styling */
    .stat-card {
        background: linear-gradient(135deg, #f8f9fa 0%, #e9ecef 100%);
        border: 1px solid #dee2e6;
        border-radius: 10px;
        padding: 1.5rem;
        text-align: center;
        box-shadow: 0 2px 8px rgba(0,0,0,0.05);
        transition: transform 0.2s ease, box-shadow 0.2s ease;
    }
    
    .stat-card:hover {
        transform: translateY(-2px);
        box-shadow: 0 4px 12px rgba(0,0,0,0.1);
    }
    
    .stat-card h3 {
        margin: 0;
        color: #1e3a5f;
        font-size: 2rem;
        font-weight: 700;
    }
    
    .stat-card p {
        margin: 0.5rem 0 0 0;
        color: #6c757d;
        font-size: 0.9rem;
        text-transform: uppercase;
        letter-spacing: 0.5px;
    }
    
    /* Success/Info boxes */
    .success-box {
        background: linear-gradient(135deg, #d4edda 0%, #c3e6cb 100%);
        border: 1px solid #28a745;
        border-radius: 10px;
        padding: 1rem 1.5rem;
        margin: 1rem 0;
    }
    
    .info-box {
        background: linear-gradient(135deg, #cce5ff 0%, #b8daff 100%);
        border: 1px solid #007bff;
        border-radius: 10px;
        padding: 1rem 1.5rem;
        margin: 1rem 0;
    }
    
    /* Button styling */
    .stDownloadButton > button {
        background: linear-gradient(135deg, #28a745 0%, #20c997 100%);
        color: white;
        border: none;
        padding: 0.75rem 2rem;
        font-size: 1.1rem;
        font-weight: 600;
        border-radius: 8px;
        transition: all 0.3s ease;
        width: 100%;
    }
    
    .stDownloadButton > button:hover {
        background: linear-gradient(135deg, #218838 0%, #1ea87a 100%);
        box-shadow: 0 4px 12px rgba(40, 167, 69, 0.3);
    }
    
    /* File uploader styling */
    .stFileUploader > div > div {
        border: 2px dashed #1e3a5f;
        border-radius: 12px;
        padding: 2rem;
        background: #f8f9fa;
    }
    
    /* Data preview table */
    .dataframe {
        font-size: 0.85rem;
    }
    
    /* Progress bar */
    .stProgress > div > div > div > div {
        background: linear-gradient(90deg, #1e3a5f 0%, #2d5a87 100%);
    }
    
    /* Sidebar styling */
    .css-1d391kg {
        background: #f8f9fa;
    }
    
    /* Section headers */
    .section-header {
        color: #1e3a5f;
        font-size: 1.3rem;
        font-weight: 600;
        margin: 1.5rem 0 1rem 0;
        padding-bottom: 0.5rem;
        border-bottom: 2px solid #1e3a5f;
    }
    
    /* Footer */
    .footer {
        text-align: center;
        padding: 2rem;
        color: #6c757d;
        font-size: 0.85rem;
        border-top: 1px solid #dee2e6;
        margin-top: 3rem;
    }
    
    /* Hide Streamlit branding */
    #MainMenu {visibility: hidden;}
    footer {visibility: hidden;}
    
    /* Expander styling */
    .streamlit-expanderHeader {
        font-weight: 600;
        color: #1e3a5f;
    }
</style>
""", unsafe_allow_html=True)


class JSONFlattener:
    """Utility class for flattening nested JSON structures."""
    
    @staticmethod
    def flatten_dict(d: Dict, parent_key: str = '', sep: str = '_') -> Dict:
        """
        Flatten a nested dictionary.
        
        Args:
            d: Dictionary to flatten
            parent_key: Parent key prefix
            sep: Separator between nested keys
            
        Returns:
            Flattened dictionary
        """
        items = []
        for k, v in d.items():
            new_key = f"{parent_key}{sep}{k}" if parent_key else k
            
            if isinstance(v, dict):
                items.extend(JSONFlattener.flatten_dict(v, new_key, sep).items())
            elif isinstance(v, list):
                if len(v) == 0:
                    items.append((new_key, None))
                elif all(isinstance(item, dict) for item in v):
                    # List of dicts - create numbered keys for first few items
                    for i, item in enumerate(v[:3]):  # Limit to first 3 items
                        items.extend(JSONFlattener.flatten_dict(item, f"{new_key}_{i}", sep).items())
                    if len(v) > 3:
                        items.append((f"{new_key}_count", len(v)))
                else:
                    # Simple list - join as string
                    items.append((new_key, ', '.join(str(item) for item in v)))
            else:
                items.append((new_key, v))
        
        return dict(items)
    
    @staticmethod
    def extract_key_fields(record: Dict) -> Dict:
        """
        Extract key fields from a commodities record for primary sheet.
        
        Args:
            record: Full JSON record
            
        Returns:
            Dictionary with key fields
        """
        result = OrderedDict()
        
        # Template Version
        result['TemplateVersion'] = record.get('TemplateVersion', '')
        
        # Header fields
        header = record.get('Header', {})
        result['AssetClass'] = header.get('AssetClass', '')
        result['InstrumentType'] = header.get('InstrumentType', '')
        result['UseCase'] = header.get('UseCase', '')
        result['Level'] = header.get('Level', '')
        
        # Identifier fields
        identifier = record.get('Identifier', {})
        result['UPI'] = identifier.get('UPI', '')
        result['Status'] = identifier.get('Status', '')
        result['StatusReason'] = identifier.get('StatusReason', '')
        result['LastUpdateDateTime'] = identifier.get('LastUpdateDateTime', '')
        
        # Derived fields
        derived = record.get('Derived', {})
        result['ClassificationType'] = derived.get('ClassificationType', '')
        result['ShortName'] = derived.get('ShortName', '')
        result['UnderlierName'] = derived.get('UnderlierName', '')
        result['UnderlyingAssetType'] = derived.get('UnderlyingAssetType', '')
        result['CFIDeliveryType'] = derived.get('CFIDeliveryType', '')
        
        # Attributes fields
        attributes = record.get('Attributes', {})
        result['ReferenceRate'] = attributes.get('ReferenceRate', '')
        result['OtherReferenceRate'] = attributes.get('OtherReferenceRate', '')
        result['BaseProduct'] = attributes.get('BaseProduct', '')
        result['SubProduct'] = attributes.get('SubProduct', '')
        result['AdditionalSubProduct'] = attributes.get('AdditionalSubProduct', '')
        result['OtherBaseProduct'] = attributes.get('OtherBaseProduct', '')
        result['OtherSubProduct'] = attributes.get('OtherSubProduct', '')
        result['OtherAdditionalSubProduct'] = attributes.get('OtherAdditionalSubProduct', '')
        result['ReturnorPayoutTrigger'] = attributes.get('ReturnorPayoutTrigger', '')
        result['DeliveryType'] = attributes.get('DeliveryType', '')
        
        return result


class JSONParser:
    """Parser for various JSON formats."""
    
    @staticmethod
    def parse_file(content: str, filename: str) -> Tuple[List[Dict], str]:
        """
        Parse JSON content from file.
        
        Args:
            content: File content as string
            filename: Original filename
            
        Returns:
            Tuple of (list of records, format description)
        """
        records = []
        format_desc = "Unknown"
        
        # Try JSON Lines / NDJSON format first
        lines = content.strip().split('\n')
        if len(lines) > 1 or (len(lines) == 1 and lines[0].startswith('{')):
            try:
                for line in lines:
                    line = line.strip()
                    if line:
                        records.append(json.loads(line))
                format_desc = "JSON Lines (NDJSON)"
                return records, format_desc
            except json.JSONDecodeError:
                pass
        
        # Try standard JSON array
        try:
            data = json.loads(content)
            if isinstance(data, list):
                records = data
                format_desc = "JSON Array"
            elif isinstance(data, dict):
                records = [data]
                format_desc = "Single JSON Object"
            return records, format_desc
        except json.JSONDecodeError as e:
            raise ValueError(f"Unable to parse JSON: {str(e)}")
        
        return records, format_desc


class ExcelGenerator:
    """Generator for Excel files from JSON data."""
    
    @staticmethod
    def create_excel(records: List[Dict], flatten_mode: str = 'smart') -> io.BytesIO:
        """
        Create Excel file from JSON records.
        
        Args:
            records: List of JSON records
            flatten_mode: 'smart', 'full', or 'minimal'
            
        Returns:
            BytesIO buffer containing Excel file
        """
        output = io.BytesIO()
        
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            if flatten_mode == 'smart':
                # Create multiple sheets for structured data
                
                # Main sheet with key fields
                main_data = [JSONFlattener.extract_key_fields(r) for r in records]
                df_main = pd.DataFrame(main_data)
                df_main.to_excel(writer, sheet_name='Main Data', index=False)
                
                # CFI details sheet
                cfi_data = []
                for i, record in enumerate(records):
                    upi = record.get('Identifier', {}).get('UPI', f'Record_{i}')
                    cfi_list = record.get('Derived', {}).get('CFI', [])
                    for cfi in cfi_list:
                        cfi_record = {'UPI': upi}
                        cfi_record['CFI_Version'] = cfi.get('Version', '')
                        cfi_record['CFI_VersionStatus'] = cfi.get('VersionStatus', '')
                        cfi_record['CFI_Value'] = cfi.get('Value', '')
                        cfi_record['CFI_Category_Code'] = cfi.get('Category', {}).get('Code', '')
                        cfi_record['CFI_Category_Value'] = cfi.get('Category', {}).get('Value', '')
                        cfi_record['CFI_Group_Code'] = cfi.get('Group', {}).get('Code', '')
                        cfi_record['CFI_Group_Value'] = cfi.get('Group', {}).get('Value', '')
                        
                        # Attributes
                        attrs = cfi.get('Attributes', [])
                        for j, attr in enumerate(attrs):
                            cfi_record[f'Attr{j+1}_Name'] = attr.get('Name', '')
                            cfi_record[f'Attr{j+1}_Code'] = attr.get('Code', '')
                            cfi_record[f'Attr{j+1}_Value'] = attr.get('Value', '')
                        
                        cfi_data.append(cfi_record)
                
                if cfi_data:
                    df_cfi = pd.DataFrame(cfi_data)
                    df_cfi.to_excel(writer, sheet_name='CFI Details', index=False)
                
                # Full flattened data sheet
                flat_data = [JSONFlattener.flatten_dict(r) for r in records]
                df_flat = pd.DataFrame(flat_data)
                df_flat.to_excel(writer, sheet_name='Full Data (Flattened)', index=False)
                
            elif flatten_mode == 'full':
                # Single sheet with fully flattened data
                flat_data = [JSONFlattener.flatten_dict(r) for r in records]
                df = pd.DataFrame(flat_data)
                df.to_excel(writer, sheet_name='Data', index=False)
                
            else:  # minimal
                # Convert to DataFrame with JSON strings for complex fields
                simple_data = []
                for record in records:
                    simple_record = {}
                    for key, value in record.items():
                        if isinstance(value, (dict, list)):
                            simple_record[key] = json.dumps(value, ensure_ascii=False)
                        else:
                            simple_record[key] = value
                    simple_data.append(simple_record)
                
                df = pd.DataFrame(simple_data)
                df.to_excel(writer, sheet_name='Data', index=False)
        
        output.seek(0)
        return output


def display_header():
    """Display the main application header."""
    st.markdown("""
    <div class="main-header">
        <h1>📊 JSON to Excel Converter</h1>
        <p>Transform your JSON data into professionally formatted Excel spreadsheets</p>
    </div>
    """, unsafe_allow_html=True)


def display_statistics(records: List[Dict], format_desc: str, filename: str):
    """Display data statistics in cards."""
    col1, col2, col3, col4 = st.columns(4)
    
    with col1:
        st.markdown(f"""
        <div class="stat-card">
            <h3>{len(records):,}</h3>
            <p>Total Records</p>
        </div>
        """, unsafe_allow_html=True)
    
    with col2:
        # Count unique fields
        all_keys = set()
        for r in records[:100]:  # Sample first 100 for performance
            all_keys.update(JSONFlattener.flatten_dict(r).keys())
        st.markdown(f"""
        <div class="stat-card">
            <h3>{len(all_keys):,}</h3>
            <p>Unique Fields</p>
        </div>
        """, unsafe_allow_html=True)
    
    with col3:
        st.markdown(f"""
        <div class="stat-card">
            <h3>{format_desc}</h3>
            <p>Source Format</p>
        </div>
        """, unsafe_allow_html=True)
    
    with col4:
        # File size estimate
        sample_size = len(json.dumps(records[0])) if records else 0
        est_size = (sample_size * len(records)) / 1024  # KB
        size_str = f"{est_size:.1f} KB" if est_size < 1024 else f"{est_size/1024:.1f} MB"
        st.markdown(f"""
        <div class="stat-card">
            <h3>{size_str}</h3>
            <p>Est. Data Size</p>
        </div>
        """, unsafe_allow_html=True)


def display_data_preview(records: List[Dict]):
    """Display data preview section."""
    st.markdown('<div class="section-header">📋 Data Preview</div>', unsafe_allow_html=True)
    
    # Show sample of key fields
    preview_data = [JSONFlattener.extract_key_fields(r) for r in records[:10]]
    df_preview = pd.DataFrame(preview_data)
    
    st.dataframe(
        df_preview,
        use_container_width=True,
        height=300
    )
    
    if len(records) > 10:
        st.caption(f"Showing first 10 of {len(records):,} records")


def display_structure_analysis(records: List[Dict]):
    """Display JSON structure analysis."""
    with st.expander("🔍 JSON Structure Analysis", expanded=False):
        if records:
            sample = records[0]
            
            col1, col2 = st.columns(2)
            
            with col1:
                st.markdown("**Top-Level Keys:**")
                for key in sample.keys():
                    value_type = type(sample[key]).__name__
                    if isinstance(sample[key], dict):
                        nested_keys = len(sample[key])
                        st.code(f"{key}: dict ({nested_keys} keys)")
                    elif isinstance(sample[key], list):
                        list_len = len(sample[key])
                        st.code(f"{key}: list ({list_len} items)")
                    else:
                        st.code(f"{key}: {value_type}")
            
            with col2:
                st.markdown("**Sample Record (Raw JSON):**")
                st.json(sample)


def main():
    """Main application function."""
    display_header()
    
    # Sidebar configuration
    with st.sidebar:
        st.markdown("### ⚙️ Configuration")
        
        st.markdown("---")
        
        flatten_mode = st.selectbox(
            "Excel Output Mode",
            options=['smart', 'full', 'minimal'],
            format_func=lambda x: {
                'smart': '🎯 Smart (Multiple Sheets)',
                'full': '📄 Full Flatten (Single Sheet)',
                'minimal': '📦 Minimal (Keep JSON Strings)'
            }[x],
            help="""
            **Smart**: Creates multiple sheets - Main Data, CFI Details, and Full Flattened
            
            **Full**: Single sheet with all nested fields flattened
            
            **Minimal**: Keeps complex fields as JSON strings
            """
        )
        
        st.markdown("---")
        
        st.markdown("### 📁 Supported Formats")
        st.markdown("""
        - `.json` - Standard JSON
        - `.ndjson` - Newline Delimited JSON
        - `.jsonl` - JSON Lines
        - `.records` - Record files
        """)
        
        st.markdown("---")
        
        st.markdown("### 💡 Tips")
        st.info("""
        • Large files may take a moment to process
        • Smart mode is recommended for complex nested data
        • Preview your data before downloading
        """)
    
    # Main content area
    st.markdown('<div class="section-header">📤 Upload Your File</div>', unsafe_allow_html=True)
    
    uploaded_file = st.file_uploader(
        "Drag and drop your JSON file here or click to browse",
        type=['json', 'ndjson', 'jsonl', 'records', 'txt'],
        help="Supports JSON, NDJSON, JSON Lines, and .records files"
    )
    
    if uploaded_file is not None:
        try:
            # Read and parse file
            with st.spinner("🔄 Parsing JSON data..."):
                content = uploaded_file.read().decode('utf-8')
                records, format_desc = JSONParser.parse_file(content, uploaded_file.name)
            
            if not records:
                st.error("❌ No valid records found in the file.")
                return
            
            # Display success message
            st.markdown(f"""
            <div class="success-box">
                ✅ <strong>Successfully parsed {len(records):,} records</strong> from {uploaded_file.name}
            </div>
            """, unsafe_allow_html=True)
            
            # Display statistics
            display_statistics(records, format_desc, uploaded_file.name)
            
            st.markdown("<br>", unsafe_allow_html=True)
            
            # Display data preview
            display_data_preview(records)
            
            # Display structure analysis
            display_structure_analysis(records)
            
            # Generate Excel section
            st.markdown('<div class="section-header">📥 Download Excel</div>', unsafe_allow_html=True)
            
            col1, col2, col3 = st.columns([1, 2, 1])
            
            with col2:
                with st.spinner("🔄 Generating Excel file..."):
                    excel_buffer = ExcelGenerator.create_excel(records, flatten_mode)
                    
                    # Generate filename
                    original_name = Path(uploaded_file.name).stem
                    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                    output_filename = f"{original_name}_converted_{timestamp}.xlsx"
                    
                    st.download_button(
                        label="⬇️ Download Excel File",
                        data=excel_buffer,
                        file_name=output_filename,
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True
                    )
                
                st.markdown(f"""
                <div class="info-box">
                    📊 <strong>Output Mode:</strong> {flatten_mode.capitalize()}<br>
                    📁 <strong>File Name:</strong> {output_filename}
                </div>
                """, unsafe_allow_html=True)
        
        except ValueError as e:
            st.error(f"❌ Error parsing file: {str(e)}")
        except Exception as e:
            st.error(f"❌ An unexpected error occurred: {str(e)}")
            st.exception(e)
    
    else:
        # Show placeholder when no file is uploaded
        st.markdown("""
        <div class="info-box">
            👆 <strong>Upload a JSON file to get started</strong><br>
            The converter supports complex nested JSON structures and will automatically
            flatten them into a clean Excel format.
        </div>
        """, unsafe_allow_html=True)
        
        # Show example
        with st.expander("📝 Example: Supported JSON Formats", expanded=False):
            st.markdown("**JSON Lines (NDJSON) - One JSON object per line:**")
            st.code('''{"id": 1, "name": "Item 1", "details": {"price": 100}}
{"id": 2, "name": "Item 2", "details": {"price": 200}}''', language='json')
            
            st.markdown("**JSON Array:**")
            st.code('''[
  {"id": 1, "name": "Item 1"},
  {"id": 2, "name": "Item 2"}
]''', language='json')
    
    # Footer
    st.markdown("""
    <div class="footer">
        <p>JSON to Excel Converter v1.0.0 | Built with Streamlit</p>
        <p>Supports complex nested JSON structures with smart flattening</p>
    </div>
    """, unsafe_allow_html=True)


if __name__ == "__main__":
    main()
