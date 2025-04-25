"""
Brief Extractor Module

This module provides functions to extract structured data from campaign briefs.
The main function breaks down a brief into logical sections:
- Account level data
- Campaign level data 
- Placement level data
- Target level data
"""

import pandas as pd
from datetime import datetime
import re
import os
from openpyxl.utils import get_column_letter

def extract_structured_brief_data(brief_path):
    """
    Extract structured data from a brief Excel file.
    
    Args:
        brief_path (str): Path to the brief Excel file
        
    Returns:
        dict: Dictionary containing structured data sections
    """
    # Initialize structured data dictionary with None values
    structured_data = {
        'account_data': None,       # DataFrame with account-level data
        'campaign_data': None,      # DataFrame with campaign-level data (including measurement/viewability)
        'placement_data': None,     # DataFrame with placement-level data
        'target_data': None         # DataFrame with target audience data
    }
    
    try:
        # Read the Excel file
        brief_df = pd.read_excel(brief_path, header=None)
        
        # Extract account-level data
        account_data = extract_account_data_from_excel(brief_df)
        if account_data:
            structured_data['account_data'] = pd.DataFrame([account_data])
        
        # Extract campaign-level data (including measurement data)
        campaign_data = extract_campaign_data_from_excel(brief_df)
        if campaign_data is not None:
            structured_data['campaign_data'] = campaign_data
        
        # Extract placement-level data
        placement_data = extract_placement_data_from_excel(brief_df)
        if placement_data:
            structured_data['placement_data'] = pd.DataFrame(placement_data)
            
            # Standardize date formats
            date_columns = ['Start Date', 'End Date']
            for col in date_columns:
                if col in structured_data['placement_data'].columns:
                    structured_data['placement_data'][col] = structured_data['placement_data'][col].apply(standardize_date_format)
        
        # Extract target audience data
        target_data = extract_target_data_from_excel(brief_df)
        if target_data:
            structured_data['target_data'] = pd.DataFrame(target_data)
        
        return structured_data
        
    except Exception as e:
        print(f"Error extracting structured data: {str(e)}")
        return structured_data

def standardize_date_format(date_str):
    """
    Standardize different date formats to MM/DD/YYYY format.
    
    Args:
        date_str (str): Date string in various formats
        
    Returns:
        str: Standardized date string in MM/DD/YYYY format
    """
    if not date_str or pd.isna(date_str):
        return date_str
        
    try:
        # If it's already a datetime object
        if isinstance(date_str, (datetime, pd.Timestamp)):
            return date_str.strftime('%m/%d/%Y')
            
        # If it's a float or integer, it might be an Excel date
        if isinstance(date_str, (int, float)):
            try:
                if 30000 < date_str < 70000:  # Reasonable Excel date range
                    date_obj = pd.to_datetime('1899-12-30') + pd.Timedelta(days=float(date_str))
                    return date_obj.strftime('%m/%d/%Y')
            except:
                pass
        
        # Try parsing with various formats
        date_formats = [
            '%Y-%m-%d', '%m/%d/%Y', '%m-%d-%Y', '%d-%m-%Y', '%d/%m/%Y',
            '%B %d, %Y', '%b %d, %Y', '%Y/%m/%d',
            '%m/%d/%y', '%d/%m/%y', '%y-%m-%d',
            '%m.%d.%Y', '%d.%m.%Y', '%Y.%m.%d'
        ]
        
        for fmt in date_formats:
            try:
                date_obj = datetime.strptime(str(date_str).strip(), fmt)
                return date_obj.strftime('%m/%d/%Y')
            except ValueError:
                continue
                
        # If no formats match, try pandas to_datetime as last resort
        try:
            date_obj = pd.to_datetime(date_str)
            return date_obj.strftime('%m/%d/%Y')
        except:
            pass
                
        # Return original if all attempts fail
        return str(date_str).strip()
    except Exception:
        return str(date_str).strip()

def extract_product_data(brief_df, data_dict):
    """
    Extract product type data from the brief
    """
    product_found = False
    
    # Method 1: Look for the Products header or Product Type section
    product_header_idx = None
    
    for idx, row in brief_df.iterrows():
        for col_idx, val in enumerate(row):
            if pd.notna(val) and isinstance(val, str) and (
                'products' in val.lower() or 'product type' in val.lower()
            ):
                product_header_idx = idx
                print(f"Found product section at row {idx+1}, col {col_idx+1}: {val}")
                break
        
        if product_header_idx is not None:
            break
    
    if product_header_idx is not None:
        # Extract rows until we find an empty row
        end_idx = product_header_idx + 1
        for idx in range(product_header_idx + 1, min(product_header_idx + 10, len(brief_df))):
            row = brief_df.iloc[idx]
            if row.isna().all():
                end_idx = idx
                break
        
        # Extract the product data
        product_data = brief_df.iloc[product_header_idx:end_idx].copy()
        
        # Check if this is a table format with headers in the first row
        if len(product_data) >= 2:
            # Use first row as headers
            headers = []
            for i, header in enumerate(product_data.iloc[0]):
                if pd.notna(header) and str(header).strip():
                    headers.append(str(header).strip())
                else:
                    headers.append(f"Column_{i}")
            
            # Extract data rows
            product_rows = product_data.iloc[1:]
            product_rows.columns = headers
            
            # Remove empty rows and columns
            product_rows = product_rows.dropna(how='all')
            
            if not product_rows.empty:
                # Store in the provided dictionary
                data_dict['product_data'] = product_rows
                product_found = True
                print(f"Found product data: {len(product_rows)} rows")
            else:
                # If table format didn't work, try key-value format
                key_value_data = extract_key_value_format(product_data)
                if key_value_data is not None:
                    data_dict['product_data'] = key_value_data
                    product_found = True
                    print(f"Found product data in key-value format: {len(key_value_data)} rows")
    
    # Method 2: If product section not found, scan the entire brief for Product Type
    if not product_found:
        print("Searching for Product Type information in the entire brief...")
        
        # Create a field-value DataFrame for product info
        product_fields = []
        product_values = []
        
        # Look for Product Type or similar in any cell
        for idx, row in brief_df.iterrows():
            for col_idx, val in enumerate(row):
                if pd.notna(val) and isinstance(val, str) and 'product type' in val.lower():
                    # Found a product type cell
                    product_fields.append('Product Type')
                    
                    # Look for the value in the next column
                    if col_idx + 1 < len(row) and pd.notna(row[col_idx + 1]):
                        product_values.append(row[col_idx + 1])
                        print(f"Found Product Type: {row[col_idx + 1]} at row {idx+1}, col {col_idx+2}")
                    else:
                        # If there's no value in the next column, check if there are alternative patterns
                        # like "Product Type: BV-Standard"
                        match = re.search(r'product type[:\s]+([^:]+)', val.lower())
                        if match:
                            product_value = match.group(1).strip()
                            product_values.append(product_value)
                            print(f"Found Product Type: {product_value} (embedded) at row {idx+1}, col {col_idx+1}")
                        else:
                            product_values.append("")
                            
                    product_found = True
                    break
            
            # Also look for "BV - Standard" or similar which often indicates product type
            if not product_found:
                for col_idx, val in enumerate(row):
                    if pd.notna(val) and isinstance(val, str) and 'bv' in val.lower() and 'standard' in val.lower():
                        product_fields.append('Product Type')
                        product_values.append(val)
                        print(f"Found Product Type: {val} at row {idx+1}, col {col_idx+1}")
                        product_found = True
                        break
        
        if product_fields and product_values:
            # Create a DataFrame with the found product info
            product_df = pd.DataFrame({
                'Field': product_fields,
                'Value': product_values
            })
            data_dict['product_data'] = product_df
            print(f"Extracted {len(product_df)} product type entries")
    
    return product_found

def extract_measurement_data(brief_df, structured_data):
    """
    Extract measurement and viewability data from the brief 
    and store it as campaign data
    """
    # Look for Measurement or Viewability headers
    measurement_header_idx = None
    
    # First attempt: Look for explicit Measurement or Viewability section headers
    for idx, row in brief_df.iterrows():
        for col_idx, val in enumerate(row):
            if pd.notna(val) and isinstance(val, str) and (
                'measurement' in val.lower() or 'viewability' in val.lower() 
            ):
                measurement_header_idx = idx
                print(f"Found measurement/viewability section at row {idx+1}, col {col_idx+1}: {val}")
                break
        
        if measurement_header_idx is not None:
            break
    
    # Second attempt: If not found explicitly, look for measurement-related terms
    if measurement_header_idx is None:
        for idx, row in brief_df.iterrows():
            for col_idx, val in enumerate(row):
                if pd.notna(val) and isinstance(val, str) and (
                    'moat' in val.lower() or 'ias' in val.lower() or 
                    'goal' in val.lower() or 'target' in val.lower() and 'viewability' in val.lower()
                ):
                    measurement_header_idx = idx
                    print(f"Found implicit measurement/viewability section at row {idx+1}, col {col_idx+1}: {val}")
                    break
            
            if measurement_header_idx is not None:
                break
    
    if measurement_header_idx is not None:
        # Find where the section ends (usually a few rows)
        end_idx = measurement_header_idx + 1
        for idx in range(measurement_header_idx + 1, min(measurement_header_idx + 15, len(brief_df))):
            row = brief_df.iloc[idx]
            # Stop if we hit a completely empty row or a new section header
            if row.isna().all() or any(
                pd.notna(val) and isinstance(val, str) and (
                    'placement' in str(val).lower() or 'target' in str(val).lower() or
                    'bv id' in str(val).lower() or 'product' in str(val).lower()
                ) for val in row
            ):
                end_idx = idx
                break
        
        # Extract measurement data section
        measurement_data = brief_df.iloc[measurement_header_idx:end_idx].copy()
        
        # Process using table format
        processed_table = process_table_format(measurement_data)
        if processed_table is not None and not processed_table.empty:
            # Store as measurement data and merge with existing campaign data if available
            if structured_data['campaign_data'] is not None:
                # If campaign data exists and both are in key-value format, combine them
                if 'Field' in processed_table.columns and 'Value' in processed_table.columns and \
                   'Field' in structured_data['campaign_data'].columns and 'Value' in structured_data['campaign_data'].columns:
                    structured_data['campaign_data'] = pd.concat([structured_data['campaign_data'], processed_table], ignore_index=True)
                else:
                    # Keep both datasets separately
                    structured_data['measurement_data'] = processed_table
            else:
                structured_data['campaign_data'] = processed_table
            
            print(f"Found measurement/viewability data: {len(processed_table)} rows")
        else:
            # Fallback to key-value extraction
            key_value_data = extract_key_value_format(measurement_data)
            if key_value_data is not None:
                # Store as measurement data and merge with existing campaign data if available
                if structured_data['campaign_data'] is not None:
                    # If both are in key-value format, combine them
                    if 'Field' in key_value_data.columns and 'Value' in key_value_data.columns and \
                       'Field' in structured_data['campaign_data'].columns and 'Value' in structured_data['campaign_data'].columns:
                        structured_data['campaign_data'] = pd.concat([structured_data['campaign_data'], key_value_data], ignore_index=True)
                    else:
                        # Keep both datasets separately
                        structured_data['measurement_data'] = key_value_data
                else:
                    structured_data['campaign_data'] = key_value_data
                
                print(f"Found measurement data in key-value format: {len(key_value_data)} rows")

def process_table_format(data_df):
    """
    Process a table format with headers in the first row and data in subsequent rows
    
    Args:
        data_df: DataFrame containing the table
        
    Returns:
        Processed DataFrame with proper headers and data rows
    """
    if len(data_df) < 2:
        return None
    
    # Get headers from first row, handling merged cells and empty headers
    header_row = data_df.iloc[0]
    headers = []
    last_header = None
    
    for i, header in enumerate(header_row):
        if pd.notna(header) and str(header).strip():
            last_header = str(header).strip()
            headers.append(last_header)
        elif last_header and i > 0:
            # Likely a merged cell - use previous header with suffix
            headers.append(f"{last_header}_{i-len(headers)+1}")
        else:
            headers.append(f"Column_{i}")
    
    # Extract data rows (skip header)
    data_rows = data_df.iloc[1:]
    data_rows.columns = headers
    
    # Remove empty rows
    data_rows = data_rows.dropna(how='all')
    
    return data_rows

def extract_key_value_format(data_df):
    """
    Extract data in a key-value format (field in one column, value in another)
    
    Args:
        data_df: DataFrame containing key-value pairs
        
    Returns:
        DataFrame with 'Field' and 'Value' columns
    """
    if data_df.empty:
        return None
    
    field_values = []
    values = []
    
    for _, row in data_df.iterrows():
        field_col = None
        value_col = None
        
        # Look for non-empty cells
        for i, cell in enumerate(row):
            if pd.notna(cell) and str(cell).strip():
                if field_col is None:
                    field_col = i
                    field_values.append(str(cell).strip())
                elif value_col is None:
                    value_col = i
                    values.append(str(cell).strip())
                    break
        
        # If we found a field but no value, add empty value
        if field_col is not None and value_col is None and len(field_values) > len(values):
            values.append("")
    
    # Create dataframe if we found any key-value pairs
    if field_values and len(field_values) == len(values):
        return pd.DataFrame({
            'Field': field_values,
            'Value': values
        })
    
    return None

def get_field_value(df, field_pattern):
    """
    Extract value for a specific field from a dataframe with Field/Value columns.
    Returns the value if found, None otherwise.
    """
    if df is None or df.empty:
        return None
        
    # Find rows matching the field pattern
    matching_rows = None
    
    # Handle potential string accessor issues
    try:
        matching_rows = df[df['Field'].str.contains(field_pattern, case=False, na=False)]
    except AttributeError:
        # Manual approach if str accessor fails
        matching_indices = []
        for idx, val in df['Field'].items():
            if pd.notna(val) and isinstance(val, str) and field_pattern.lower() in val.lower():
                matching_indices.append(idx)
        
        if matching_indices:
            matching_rows = df.loc[matching_indices]
    
    if matching_rows is not None and not matching_rows.empty:
        # Return the value from the first matching row
        return matching_rows.iloc[0]['Value']
    
    return None

def export_to_excel(structured_data, output_path):
    """
    Export structured brief data to an Excel file with separate tabs for each section.
    
    Args:
        structured_data (dict): Dictionary containing structured brief data
        output_path (str): Path to save the Excel file
    """
    print(f"Exporting structured data to {output_path}")
    
    # Create a Pandas Excel writer using the specified filename and engine
    writer = pd.ExcelWriter(output_path, engine='openpyxl')
    
    # Set column width function
    def set_column_width(worksheet):
        for column in worksheet.columns:
            max_length = 0
            column_letter = get_column_letter(column[0].column)
            for cell in column:
                try:
                    if len(str(cell.value)) > max_length:
                        max_length = len(str(cell.value))
                except:
                    pass
            adjusted_width = (max_length + 2)
            worksheet.column_dimensions[column_letter].width = adjusted_width
    
    # Create a default empty DataFrame for the first sheet
    default_df = pd.DataFrame({'Note': ['No data found in the brief.']})
    
    # Write each DataFrame to a separate sheet with auto-column width
    sheets_written = False
    
    if structured_data['account_data'] is not None and not structured_data['account_data'].empty:
        structured_data['account_data'].to_excel(writer, sheet_name='Account Level', index=False)
        set_column_width(writer.sheets['Account Level'])
        sheets_written = True
    
    if structured_data['campaign_data'] is not None and not structured_data['campaign_data'].empty:
        structured_data['campaign_data'].to_excel(writer, sheet_name='Campaign Level', index=False)
        set_column_width(writer.sheets['Campaign Level'])
        sheets_written = True
    
    if structured_data['placement_data'] is not None and not structured_data['placement_data'].empty:
        structured_data['placement_data'].to_excel(writer, sheet_name='Placement Level', index=False)
        set_column_width(writer.sheets['Placement Level'])
        sheets_written = True
        
    if structured_data['target_data'] is not None and not structured_data['target_data'].empty:
        structured_data['target_data'].to_excel(writer, sheet_name='Target Level', index=False)
        set_column_width(writer.sheets['Target Level'])
        sheets_written = True
    
    # If no sheets were written, create a default sheet
    if not sheets_written:
        default_df.to_excel(writer, sheet_name='No Data Found', index=False)
        set_column_width(writer.sheets['No Data Found'])
    
    # Save and close the workbook
    writer.close()
    print(f"Structured brief data exported to {output_path}")
    
    return output_path

def test_extraction():
    """Test function for brief extraction"""
    brief_path = "C:/QA_auto_check/Brief/Campaign_Brief.xlsx"
    output_path = "structured_brief.xlsx"  # Save in current directory
    
    if not os.path.exists(brief_path):
        print(f"Error: Brief file not found at {brief_path}")
        return
    
    print(f"Testing structured brief extraction on: {brief_path}")
    
    try:
        # Extract structured data
        structured_brief = extract_structured_brief_data(brief_path)
        
        if structured_brief:
            # Print all data sections in a combined summary
            print("\n=== STRUCTURED BRIEF DATA SUMMARY ===")
            
            # Print account data
            if structured_brief['account_data'] is not None and not structured_brief['account_data'].empty:
                print("\n--- Account Level Data ---")
                print(f"Found {len(structured_brief['account_data'])} rows of account data")
                for _, row in structured_brief['account_data'].iterrows():
                    for col in structured_brief['account_data'].columns:
                        if pd.notna(row[col]):
                            print(f"  {col}: {row[col]}")
            else:
                print("\n--- Account Level Data: None Found ---")
            
            # Print campaign data
            if structured_brief['campaign_data'] is not None and not structured_brief['campaign_data'].empty:
                print("\n--- Campaign Level Data ---")
                print(f"Found {len(structured_brief['campaign_data'])} rows of campaign data")
                for _, row in structured_brief['campaign_data'].iterrows():
                    for col in structured_brief['campaign_data'].columns:
                        if pd.notna(row[col]):
                            print(f"  {col}: {row[col]}")
            else:
                print("\n--- Campaign Level Data: None Found ---")
            
            # Print placement data
            if structured_brief['placement_data'] is not None and not structured_brief['placement_data'].empty:
                print("\n--- Placement Level Data ---")
                print(f"Found {len(structured_brief['placement_data'])} placements")
                for idx, placement in structured_brief['placement_data'].iterrows():
                    print(f"\nPlacement {idx + 1}:")
                    for col in structured_brief['placement_data'].columns:
                        if pd.notna(placement[col]):
                            print(f"  {col}: {placement[col]}")
            else:
                print("\n--- Placement Level Data: None Found ---")
            
            # Print target data
            if structured_brief['target_data'] is not None and not structured_brief['target_data'].empty:
                print("\n--- Target Level Data ---")
                print(f"Found {len(structured_brief['target_data'])} targets")
                for idx, target in structured_brief['target_data'].iterrows():
                    print(f"\nTarget {idx + 1}:")
                    for col in structured_brief['target_data'].columns:
                        if pd.notna(target[col]):
                            print(f"  {col}: {target[col]}")
            else:
                print("\n--- Target Level Data: None Found ---")
            
            # Export structured data to Excel
            export_to_excel(structured_brief, output_path)
            print(f"\nStructured brief data has been exported to {output_path}")
        else:
            print("No structured data could be extracted from the brief.")
    
    except Exception as e:
        print(f"Error during extraction: {str(e)}")
        # Create a default Excel file with error message
        error_df = pd.DataFrame({'Error': [f'Error during extraction: {str(e)}']})
        error_df.to_excel(output_path, sheet_name='Error', index=False)
        print(f"Error details have been saved to {output_path}")

def extract_campaign_data(brief_text):
    """
    Extract campaign-level data from the brief text.
    
    Args:
        brief_text (str): The brief text content
        
    Returns:
        dict: Dictionary containing campaign-level data
    """
    campaign_data = {}
    
    # Define patterns for campaign data
    campaign_patterns = {
        'IO Campaign Start Date': r'IO Campaign Start Date[:\s]+([^\n]+)',
        'IO Campaign End Date': r'IO Campaign End Date[:\s]+([^\n]+)',
        'BV Budget': r'BV Budget[:\s]+([^\n]+)',
        'Apply Blacklist or Whitelist': r'Apply Blacklist or Whitelist[:\s]+([^\n]+)',
        'Exclusion or Inclusion List Notes': r'Exclusion or Inclusion List Notes[:\s]+([^\n]+)',
        'Apply Dairy-Milk Restrictions': r'Apply Dairy-Milk Restrictions[:\s]+([^\n]+)',
        'LDA or Age Compliant': r'LDA or Age Compliant[:\s]+([^\n]+)',
        # Measurement and viewability patterns
        'Measurement Type': r'Measurement Type[:\s]+([^\n]+)',
        'Viewability Contracted': r'Viewability Contracted[:\s]+([^\n]+)',
        'Viewability Goal': r'Viewability Goal[:\s]+([^\n]+)',
    }
    
    # Extract data using patterns
    for field, pattern in campaign_patterns.items():
        match = re.search(pattern, brief_text, re.IGNORECASE)
        if match:
            campaign_data[field] = match.group(1).strip()
    
    return campaign_data

def extract_account_data(brief_text):
    """
    Extract account-level data from the brief text.
    
    Args:
        brief_text (str): The brief text content
        
    Returns:
        dict: Dictionary containing account-level data
    """
    account_data = {}
    
    # Define patterns for account data
    account_patterns = {
        "Today's Date": r"Today's Date[:\s]+([^\n]+)",
        "Account Name": r"Account Name[:\s]+([^\n]+)",
        "Campaign Name": r"Campaign Name[:\s]+([^\n]+)",
        "Business Consultant": r"Business Consultant[:\s]+([^\n]+)",
        "Campaign Specialist": r"Campaign Specialist[:\s]+([^\n]+)",
        "Business Account Manager": r"Business Account Manager[:\s]+([^\n]+)",
        "Ad Ops Specialist": r"Ad Ops Specialist[:\s]+([^\n]+)",
        "Product Type": r"Product Type[:\s]+([^\n]+)"
    }
    
    # Extract data using patterns
    for field, pattern in account_patterns.items():
        match = re.search(pattern, brief_text, re.IGNORECASE)
        if match:
            account_data[field] = match.group(1).strip()
    
    return account_data

def extract_placement_data(brief_text):
    """
    Extract placement-level data from the brief text.
    
    Args:
        brief_text (str): The brief text content
        
    Returns:
        list: List of dictionaries containing placement data
    """
    placements = []
    
    # Find the placement section
    placement_section = re.search(r'Placement Data(.*?)(?=Target Data|$)', brief_text, re.DOTALL | re.IGNORECASE)
    if not placement_section:
        return placements
    
    placement_text = placement_section.group(1)
    
    # Split into individual placements
    placement_blocks = re.split(r'\n\s*\n', placement_text)
    
    for block in placement_blocks:
        if not block.strip():
            continue
            
        placement = {}
        
        # Extract placement fields
        placement_fields = {
            'BV Placement Name': r'BV Placement Name[:\s]+([^\n]+)',
            'BVP': r'BVP[:\s]+([^\n]+)',
            'Start Date': r'Start Date[:\s]+([^\n]+)',
            'End Date': r'End Date[:\s]+([^\n]+)',
            'Platform/Media Type': r'Platform/Media Type[:\s]+([^\n]+)',
            'Geo Required': r'Geo Required[:\s]+([^\n]+)',
            'Budget': r'Budget[:\s]+([^\n]+)'
        }
        
        for field, pattern in placement_fields.items():
            match = re.search(pattern, block, re.IGNORECASE)
            if match:
                placement[field] = match.group(1).strip()
        
        if placement:  # Only add if we found any data
            placements.append(placement)
    
    return placements

def extract_target_data(brief_text):
    """
    Extract target audience data from the brief text.
    
    Args:
        brief_text (str): The brief text content
        
    Returns:
        list: List of dictionaries containing target data
    """
    targets = []
    
    # Find the target section
    target_section = re.search(r'Target Data(.*?)(?=Placement Data|$)', brief_text, re.DOTALL | re.IGNORECASE)
    if not target_section:
        return targets
    
    target_text = target_section.group(1)
    
    # Split into individual targets
    target_blocks = re.split(r'\n\s*\n', target_text)
    
    for block in target_blocks:
        if not block.strip():
            continue
            
        target = {}
        
        # Extract target fields
        target_fields = {
            'BV ID': r'BV ID[:\s]+([^\n]+)',
            'BVP': r'BVP[:\s]+([^\n]+)',
            'BVT': r'BVT[:\s]+([^\n]+)',
            'Target Description': r'Target Description[:\s]+([^\n]+)',
            'Target Type': r'Target Type[:\s]+([^\n]+)',
            'Target Value': r'Target Value[:\s]+([^\n]+)'
        }
        
        for field, pattern in target_fields.items():
            match = re.search(pattern, block, re.IGNORECASE)
            if match:
                target[field] = match.group(1).strip()
        
        if target:  # Only add if we found any data
            targets.append(target)
    
    return targets

def extract_account_data_from_excel(brief_df):
    """
    Extract account-level data from the Excel brief.
    
    Args:
        brief_df (DataFrame): The brief Excel data
        
    Returns:
        dict: Dictionary containing account-level data
    """
    account_data = {}
    
    # Look for account data in first few rows
    for idx, row in brief_df.iloc[0:30].iterrows(): # Increased range to 30 rows
        for col_idx, val in enumerate(row):
            if pd.notna(val) and isinstance(val, str):
                # Check for account fields
                for field in ["Today's Date", "Account Name", "Campaign Name", 
                            "Business Consultant", "Campaign Specialist", "Business Account Manager", 
                            "Ad Ops Specialist", "Product Type"]:
                    if field.lower() in val.lower():
                        # Get the value from the next column or the one after
                        value = None
                        # Check next column
                        if col_idx + 1 < len(row) and pd.notna(row[col_idx + 1]):
                            value = str(row[col_idx + 1]).strip()
                        # If not found or empty, check the column after that
                        elif col_idx + 2 < len(row) and pd.notna(row[col_idx + 2]):
                             value = str(row[col_idx + 2]).strip()

                        if value: # Only add if a non-empty value was found
                             account_data[field] = value
                        break # Move to the next cell once a field is matched
    
    return account_data

def extract_campaign_data_from_excel(brief_df):
    """
    Extract campaign-level data from the Excel brief.
    
    Args:
        brief_df (DataFrame): The brief Excel data
        
    Returns:
        DataFrame: DataFrame containing campaign-level data in Field/Value format
    """
    # Initialize lists to store field-value pairs
    fields = []
    values = []
    
    # Define the specific fields we're looking for with exact matches
    target_fields = {
        'IO Campaign Start Date': ['io campaign start date'],
        'IO Campaign End Date': ['io campaign end date', 'io campaign  end date'],
        'Apply Dairy-Milk Restrictions': ['apply dairy-milk restrictions', 'apply dairy milk restrictions'],
        'LDA or Age Compliant': ['lda or age compliant'],
        'LDA or Age Compliant Notes': ['lda or age compliant notes'],
        'BV Budget': ['bv budget'],
        'Measurement Type': ['measurement type'],
        'Viewability Contracted': ['viewability contracted'],
        'Viewability Goal': ['viewability goal']
    }
    
    # Track which fields we've found to avoid duplicates
    found_fields = set()
    
    # Look for campaign data in rows 0-30
    for idx, row in brief_df.iloc[0:30].iterrows():
        if row.isna().all():  # Skip empty rows
            continue
            
        for col_idx, cell in enumerate(row):
            if pd.isna(cell):  # Skip empty cells
                continue
                
            cell_text = str(cell).strip().lower()
            
            # Check each field
            for field, variations in target_fields.items():
                if field in found_fields:
                    continue
                    
                if any(var in cell_text for var in variations):
                    value = None
                    
                    # Check next column for value
                    if col_idx + 1 < len(row):
                        next_cell = row[col_idx + 1]
                        if pd.notna(next_cell):
                            # Handle date fields
                            if 'date' in field.lower():
                                try:
                                    date_value = pd.to_datetime(next_cell)
                                    value = date_value.strftime('%m/%d/%Y')
                                except:
                                    value = str(next_cell).strip()
                            else:
                                value = str(next_cell).strip()
                    
                    # Add field and value (empty string if no value found)
                    fields.append(field)
                    values.append(value if value else "")
                    found_fields.add(field)
                    print(f"Found {field}: {value if value else '(empty)'}")
                    break
    
    if fields and values:
        # Create DataFrame with Field/Value columns
        campaign_df = pd.DataFrame({
            'Field': fields,
            'Value': values
        })
        
        # Sort the DataFrame to match the order of target_fields
        field_order = {field: idx for idx, field in enumerate(target_fields.keys())}
        campaign_df['sort_order'] = campaign_df['Field'].map(field_order)
        campaign_df = campaign_df.sort_values('sort_order').drop('sort_order', axis=1)
        
        return campaign_df
    
    return None

def extract_placement_data_from_excel(brief_df):
    """
    Extract placement-level data from the Excel brief.
    
    Args:
        brief_df (DataFrame): The brief Excel data
        
    Returns:
        list: List of dictionaries containing placement data
    """
    placements = []
    
    # Find the placement header row
    placement_header_idx = None
    for idx, row in brief_df.iterrows():
        if idx < 20:  # Skip initial rows
            continue
        for val in row:
            if pd.notna(val) and isinstance(val, str) and ('placement name' in val.lower() or 'bvp' in val.lower()):
                placement_header_idx = idx
                break
        if placement_header_idx is not None:
            break
    
    if placement_header_idx is not None:
        # Find where placement data ends
        placement_end_idx = None
        for idx in range(placement_header_idx + 1, len(brief_df)):
            row = brief_df.iloc[idx]
            if row.isna().all() or (pd.notna(row[1]) and 'bv id' in str(row[1]).lower()):
                placement_end_idx = idx
                break
        
        if placement_end_idx:
            # Extract placement data
            placement_data = brief_df.iloc[placement_header_idx:placement_end_idx].copy()
            # Use first row as headers
            headers = [str(h) if pd.notna(h) else f'Column_{i}' for i, h in enumerate(placement_data.iloc[0])]
            placement_data = placement_data.iloc[1:]
            placement_data.columns = headers
            
            # Convert to list of dictionaries
            for _, row in placement_data.iterrows():
                placement = {}
                for col in headers:
                    if pd.notna(row[col]):
                        placement[col] = str(row[col]).strip()
                if placement:
                    placements.append(placement)
    
    return placements

def extract_target_data_from_excel(brief_df):
    """
    Extract target audience data from the Excel brief.
    
    Args:
        brief_df (DataFrame): The brief Excel data
        
    Returns:
        list: List of dictionaries containing target data
    """
    targets = []
    
    # Find the target header row
    target_header_idx = None
    for idx, row in brief_df.iterrows():
        if idx < 20:  # Skip initial rows
            continue
        if (pd.notna(row[1]) and 'bv id' in str(row[1]).lower() and
            pd.notna(row[2]) and 'bvp' in str(row[2]).lower() and
            pd.notna(row[3]) and 'bvt' in str(row[3]).lower()):
            target_header_idx = idx
            break
    
    if target_header_idx is not None:
        # Extract target data
        target_data = brief_df.iloc[target_header_idx:].copy()
        # Find where target data ends
        for i, row in target_data.iterrows():
            if row.isna().all():
                target_data = target_data.loc[:i-1]
                break
        
        # Use first row as headers
        headers = [str(h) if pd.notna(h) else f'Column_{i}' for i, h in enumerate(target_data.iloc[0])]
        target_data = target_data.iloc[1:]
        target_data.columns = headers
        
        # Convert to list of dictionaries
        for _, row in target_data.iterrows():
            target = {}
            for col in headers:
                if pd.notna(row[col]):
                    target[col] = str(row[col]).strip()
            if target:
                targets.append(target)
    
    return targets

if __name__ == "__main__":
    test_extraction()

# brief_extractor.y_4.15_11.51am