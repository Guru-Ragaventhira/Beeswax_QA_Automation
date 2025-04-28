import streamlit as st
import os
import tempfile
import shutil
from datetime import datetime
import argparse  # Import for command line arguments
import pandas as pd
import importlib.util
import sys
import time
from brief import run_qa_checks  # Import the QA checks function
from dotenv import load_dotenv  # Explicitly import dotenv

def load_module_from_file(module_name, file_path):
    """
    Dynamically load a module from a file path
    """
    spec = importlib.util.spec_from_file_location(module_name, file_path)
    module = importlib.util.module_from_spec(spec)
    sys.modules[module_name] = module
    spec.loader.exec_module(module)
    return module

def save_uploaded_file(uploaded_file, temp_dir):
    """Save uploaded file to temporary directory and return its path."""
    file_path = os.path.join(temp_dir, uploaded_file.name)
    with open(file_path, "wb") as f:
        f.write(uploaded_file.getbuffer())
    return file_path

def ensure_output_dir(output_dir="output_reports"):
    """Ensure output directory exists."""
    if not os.path.exists(output_dir):
        os.makedirs(output_dir)
    return output_dir

def load_credentials():
    """
    Load credentials securely, either from Streamlit secrets (when deployed)
    or from local .env file (when running locally)
    """
    # Initialize credentials dictionary
    credentials = {}
    
    # Try to load from Streamlit secrets first (for cloud deployment)
    try:
        if 'beeswax_credentials' in st.secrets:
            st.success("✓ Using secure credentials from Streamlit")
            # Map Streamlit secrets to environment variables
            for key, value in st.secrets.beeswax_credentials.items():
                os.environ[key] = value
                # Store masked values for critical credentials
                if key in ['LOGIN_EMAIL', 'PASSWORD']:
                    if value:
                        masked_value = value[:3] + '*' * (len(value) - 6) + value[-3:] if len(value) > 6 else "****"
                        credentials[key] = masked_value
                else:
                    credentials[key] = value
            return True, credentials
    except Exception as e:
        st.warning(f"Could not load Streamlit secrets: {e}")
    
    # If not running on Streamlit Cloud or secrets not configured, try local .env
    env_path = os.environ.get('ENV_PATH', './input_folder/beeswax_input_qa.env')
    
    if not os.path.exists(env_path):
        return False, f"Environment file not found: {env_path}"
    
    # Load environment variables
    load_dotenv(env_path, override=True)
    
    # Check if critical variables were loaded
    required_vars = ['LOGIN_EMAIL', 'PASSWORD']
    missing_vars = [var for var in required_vars if not os.getenv(var)]
    
    if missing_vars:
        return False, f"Missing required environment variables: {', '.join(missing_vars)}"
    
    # Fix specific environment variable issues
    # If LOGIN_URL exists but V2_LOGIN_URL doesn't, copy it
    if os.getenv('LOGIN_URL') and not os.getenv('V2_LOGIN_URL'):
        os.environ['V2_LOGIN_URL'] = os.getenv('LOGIN_URL')
    
    # Prepare masked credentials for display
    for key in ['LOGIN_EMAIL', 'PASSWORD', 'V2_LOGIN_URL', 'LOGIN_URL', 'CAMPAIGN_URL', 'LINEITEM_URL', 'CREATIVE_URL']:
        value = os.getenv(key)
        if key in ['LOGIN_EMAIL', 'PASSWORD']:  # Mask sensitive data
            if value:
                masked_value = value[:3] + '*' * (len(value) - 6) + value[-3:] if len(value) > 6 else "****"
                credentials[key] = masked_value
        else:
            credentials[key] = value
    
    return True, credentials

def display_ids_summary(qa_instance, title="IDs Found in Brief"):
    """Display summary of found IDs in a styled format."""
    if not qa_instance:
        return
    
    st.subheader(title)
    
    # Create columns for different ID types
    col1, col2, col3 = st.columns(3)
    
    # Display Campaign IDs
    with col1:
        st.markdown("#### Campaign IDs")
        if qa_instance.campaign_ids:
            st.success(f"Found {len(qa_instance.campaign_ids)} BVI IDs")
            for campaign_id in qa_instance.campaign_ids:
                st.write(f"- {campaign_id}")
        else:
            st.warning("No BVI IDs found")
    
    # Display Line Item IDs
    with col2:
        st.markdown("#### Line Item IDs")
        if qa_instance.line_item_ids:
            st.success(f"Found {len(qa_instance.line_item_ids)} BVT IDs")
            for line_item_id in qa_instance.line_item_ids:
                st.write(f"- {line_item_id}")
        else:
            st.warning("No BVT IDs found")
    
    # Display Creative IDs
    with col3:
        st.markdown("#### Creative IDs")
        if qa_instance.creative_ids:
            st.success(f"Found {len(qa_instance.creative_ids)} BVP IDs")
            for creative_id in qa_instance.creative_ids:
                st.write(f"- {creative_id}")
        else:
            st.warning("No BVP IDs found")

def display_fetched_ids_summary(qa_instance):
    """Display summary of successfully fetched data."""
    if not qa_instance:
        return
    
    st.subheader("Data Successfully Fetched For")
    
    # Create columns for different data types
    col1, col2, col3 = st.columns(3)
    
    # Display fetched campaign data
    with col1:
        st.markdown("#### Campaign Data")
        if qa_instance.campaign_data is not None and not qa_instance.campaign_data.empty:
            fetched_campaigns = set()
            if 'campaign_alternative_id' in qa_instance.campaign_data.columns:
                fetched_campaigns = set(qa_instance.campaign_data['campaign_alternative_id'].unique())
            elif 'campaign_id' in qa_instance.campaign_data.columns:
                fetched_campaigns = set(qa_instance.campaign_data['campaign_id'].unique())
            
            st.success(f"Fetched data for {len(fetched_campaigns)} campaigns")
            for campaign_id in fetched_campaigns:
                st.write(f"- {campaign_id}")
        else:
            st.warning("No campaign data fetched")
    
    # Display fetched line item data
    with col2:
        st.markdown("#### Line Item Data")
        if qa_instance.line_item_data is not None and not qa_instance.line_item_data.empty:
            fetched_line_items = set()
            if 'line_item_alternative_id' in qa_instance.line_item_data.columns:
                fetched_line_items = set(qa_instance.line_item_data['line_item_alternative_id'].unique())
            elif 'line_item_id' in qa_instance.line_item_data.columns:
                fetched_line_items = set(qa_instance.line_item_data['line_item_id'].unique())
            
            st.success(f"Fetched data for {len(fetched_line_items)} line items")
            for line_item_id in fetched_line_items:
                st.write(f"- {line_item_id}")
        else:
            st.warning("No line item data fetched")
    
    # Display fetched creative data
    with col3:
        st.markdown("#### Creative Data")
        if qa_instance.creative_data is not None and not qa_instance.creative_data.empty:
            fetched_creatives = set()
            if 'creative_alternative_id' in qa_instance.creative_data.columns:
                fetched_creatives = set(qa_instance.creative_data['creative_alternative_id'].unique())
            elif 'creative_id' in qa_instance.creative_data.columns:
                fetched_creatives = set(qa_instance.creative_data['creative_id'].unique())
            
            st.success(f"Fetched data for {len(fetched_creatives)} creatives")
            for creative_id in fetched_creatives:
                st.write(f"- {creative_id}")
        else:
            st.warning("No creative data fetched")

def main():
    st.title("Beeswax QA Automation")
    st.write("Upload your campaign brief to generate QA reports")
    
    # Load and validate credentials
    env_path = os.environ.get('ENV_PATH', './input_folder/beeswax_input_qa.env')
    output_dir = os.environ.get('OUTPUT_DIR', './output_folder')
    
    # Load and validate environment file silently
    success, credentials_info = load_credentials()
    
    if not success:
        st.error(f"Error loading credentials: {credentials_info}")
        st.stop()
    
    # Ensure output directory exists
    output_dir = ensure_output_dir(output_dir)
    
    # Create temporary directory for uploaded files
    with tempfile.TemporaryDirectory() as temp_dir:
        # File uploader for campaign brief only
        brief_file = st.file_uploader("Upload Campaign Brief (Excel/CSV)", type=['xlsx', 'csv'])
        
        if brief_file:
            # Save uploaded file
            brief_path = save_uploaded_file(brief_file, temp_dir)
            
            # Step 1: Run QA checks on the brief
            st.write("Running QA checks on the uploaded brief...")
            issues = []
            qa_processed_path = None
            
            try:
                issues = run_qa_checks(brief_path)
                qa_processed_path = brief_path.replace('.xlsx', '_QA_issues.xlsx')  # Path to QA-processed brief
                
                if issues:
                    st.error("❌ Issues found in the campaign brief:")
                    for issue in issues:
                        st.write(f"- {issue}")
                    st.warning("You can still generate the QA report for further analysis.")
                else:
                    st.success("✅ No issues found in the campaign brief. Proceeding with further processing.")
                
                # Display QA-processed brief and add download button in the first step
                if os.path.exists(qa_processed_path):
                    st.markdown("### QA-Processed Campaign Brief")
                    qa_brief_df = pd.read_excel(qa_processed_path)
                    st.dataframe(qa_brief_df)
                    
                    # Download button for QA-processed brief (moved to first step)
                    with open(qa_processed_path, 'rb') as f:
                        st.download_button(
                            label="Download QA-Processed Brief",
                            data=f.read(),
                            file_name=os.path.basename(qa_processed_path),
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                        )
            
            except Exception as e:
                st.error(f"An error occurred during QA checks: {str(e)}")
                st.warning("You can still generate the QA report for further analysis.")
            
            # Step 2: Display the "Generate QA Report" button
            if st.button("Generate QA Report"):
                with st.spinner("Generating Comprehensive QA Report..."):
                    try:
                        # Set environment variables for run_qa.py
                        os.environ["BRIEF_PATH"] = brief_path
                        os.environ["ENV_PATH"] = env_path
                        os.environ["OUTPUT_DIR"] = output_dir
                        # Set additional variables that run_qa.py might use
                        output_raw_dir = os.path.join(output_dir, "raw")
                        os.environ["OUTPUT_RAW_DIR"] = output_raw_dir
                        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                        combined_output_path = os.path.join(output_dir, f"combined_qa_report_{timestamp}.xlsx")
                        os.environ["COMBINED_OUTPUT_PATH"] = combined_output_path
                        
                        # Create raw output directory if it doesn't exist
                        if not os.path.exists(output_raw_dir):
                            os.makedirs(output_raw_dir)
                        
                        # First load environment variables again to ensure all scripts have the same settings
                        success, _ = load_credentials()
                        if not success:
                            st.error(f"Failed to load credentials")
                            return
                            
                        # Explicitly set V2_LOGIN_URL if needed
                        if os.getenv('LOGIN_URL') and not os.getenv('V2_LOGIN_URL'):
                            os.environ['V2_LOGIN_URL'] = os.getenv('LOGIN_URL')
                        
                        # Import and run the combined QA script
                        st.write("Loading QA modules...")
                        comb_qa = load_module_from_file("comb_qa", "run_qa.py")
                        
                        st.write("Running comprehensive QA process...")
                        final_report_path = comb_qa.main()  # This should run all QA scripts and return the combined report path
                        
                        # Allow time for file system operations to complete
                        time.sleep(1)
                        
                        if final_report_path and os.path.exists(final_report_path):
                            st.success("Comprehensive QA Report generated successfully!")
                            
                            # Load the beeswax_api module to display ID summaries
                            beeswax_api = load_module_from_file("beeswax_api", "beeswax_api.py")
                            qa = beeswax_api.BeeswaxQA(brief_path, env_path, temp_dir)
                            qa.load_brief()
                            
                            # Display ID summaries
                            display_ids_summary(qa)
                            
                            # Extract the info from comb_qa's generated report
                            excel_file = pd.ExcelFile(final_report_path)
                            sheet_names = excel_file.sheet_names
                            
                            # Create dynamic tabs based on the sheets in the Excel file
                            tabs = st.tabs(sheet_names)
                            
                            # Display each sheet in its own tab
                            for i, sheet_name in enumerate(sheet_names):
                                with tabs[i]:
                                    sheet_df = pd.read_excel(excel_file, sheet_name=sheet_name)
                                    st.dataframe(sheet_df)
                            
                            # Download button for complete QA report
                            with open(final_report_path, 'rb') as f:
                                st.download_button(
                                    label="Download Complete QA Report",
                                    data=f.read(),
                                    file_name=os.path.basename(final_report_path),
                                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                                )
                        else:
                            st.error(f"Failed to generate QA report. Please check the logs for details.")
                            st.write(f"Report path checked: {final_report_path}")
                            # List files in output directory to debug
                            if os.path.exists(output_dir):
                                files = os.listdir(output_dir)
                                st.write(f"Files in output directory ({output_dir}):")
                                for file in files:
                                    st.write(f"- {file}")
            
                    except Exception as e:
                        st.error(f"An error occurred: {str(e)}")
                        st.error("Please check the logs for more details.")

if __name__ == "__main__":
    parser = argparse.ArgumentParser(description='Beeswax QA Automation Streamlit App')
    parser.add_argument('--host', type=str, default='0.0.0.0', 
                        help='Host address to run the Streamlit app on (default: 0.0.0.0, accessible from other machines)')
    parser.add_argument('--port', type=int, default=8501, 
                        help='Port to run the Streamlit app on (default: 8501)')
    parser.add_argument('--env-path', type=str, default='./input_folder/beeswax_input_qa.env',
                        help='Path to the environment file (default: ./input_folder/beeswax_input_qa.env)')
    parser.add_argument('--output-dir', type=str, default='./output_folder',
                        help='Directory for output files (default: ./output_folder)')
    parser.add_argument('--debug', action='store_true',
                        help='Enable debug mode to show additional information')
    args = parser.parse_args()

    # Set environment variables from command line arguments
    os.environ['ENV_PATH'] = args.env_path
    os.environ['OUTPUT_DIR'] = args.output_dir
    
    # Load environment variables before starting the app
    success, env_info = load_credentials()
    
    # Print instructions for users
    print(f"\n===== Beeswax QA Automation =====")
    print(f"Starting Streamlit server on {args.host}:{args.port}")
    print(f"Using environment file: {args.env_path}")
    print(f"Using output directory: {args.output_dir}")
    print(f"Environment loaded successfully: {success}")
    
    if args.debug or not success:
        print(f"Environment details: {env_info}")
    
    print(f"To access from another machine on the same network, use: http://<your-ip-address>:{args.port}")
    print(f"Your IP address might be: {os.popen('ipconfig | findstr IPv4').read().strip()}")
    print(f"=============================================\n")

    # Pass the host and port arguments to Streamlit
    os.environ['STREAMLIT_SERVER_PORT'] = str(args.port)
    os.environ['STREAMLIT_SERVER_HEADLESS'] = 'true'
    os.environ['STREAMLIT_SERVER_ADDRESS'] = args.host
    
    main()