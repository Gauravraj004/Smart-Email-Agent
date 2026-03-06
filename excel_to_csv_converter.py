"""
Excel to CSV Converter for Cold Email Automation
Converts Tracxn Excel exports to CSV format for cold_email_automation.py
"""

import os
import re
import pandas as pd
from pathlib import Path


def extract_first_name(full_name: str) -> str:
    """Extract first name from full name, handling titles and edge cases."""
    if pd.isna(full_name) or not full_name:
        return ""
    
    name_str = str(full_name).strip()
    if not name_str or name_str.lower() in ['nan', 'none', '']:
        return ""
    
    parts = name_str.split()
    if parts:
        first = parts[0]
        # Remove common titles
        titles = ['Mr.', 'Mrs.', 'Ms.', 'Dr.', 'Prof.', 'Mr', 'Mrs', 'Ms', 'Dr', 'Prof']
        if first in titles and len(parts) > 1:
            first = parts[1]
        return first.strip('.,')
    return name_str


def is_valid_email(email: str) -> bool:
    """Check if email is valid format."""
    if pd.isna(email) or not email:
        return False
    email_str = str(email).strip().lower()
    if not email_str or email_str in ['nan', 'none', '']:
        return False
    # Basic email regex
    pattern = r'^[\w\.-]+@[\w\.-]+\.[a-zA-Z]{2,}$'
    return bool(re.match(pattern, email_str))


def process_excel_file(excel_path: str, output_dir: str) -> int:
    """
    Process a single Excel file and extract data from 'People 2.1' sheet.
    
    Args:
        excel_path: Path to the Excel file
        output_dir: Directory to save the output CSV
        
    Returns:
        Number of rows exported
    """
    filename = os.path.basename(excel_path)
    print(f"\n📂 Processing: {filename}")
    
    try:
        # Read Excel file
        xl = pd.ExcelFile(excel_path)
        
        # Check if 'People 2.1' sheet exists
        if 'People 2.1' not in xl.sheet_names:
            print(f"  ⚠️ Sheet 'People 2.1' not found. Available sheets: {xl.sheet_names}")
            return 0
        
        # Read the People 2.1 sheet (header is at row 6, index 5)
        df = pd.read_excel(excel_path, sheet_name='People 2.1', header=5)
        
        print(f"  📊 Found {len(df)} rows in People 2.1 sheet")
        
        # Extract required columns
        output_data = []
        
        for _, row in df.iterrows():
            company_name = str(row.get('Company Name', '')).strip()
            founder_name = str(row.get('Founder Name', '')).strip()
            email = str(row.get('Emails', '')).strip()
            
            # Skip rows without valid email
            if not is_valid_email(email):
                continue
            
            # Extract first name
            first_name = extract_first_name(founder_name)
            
            # Fallback for empty first name: use email username
            if not first_name or first_name.lower() in ['nan', 'none']:
                first_name = email.split('@')[0].capitalize()
            
            # Fallback for empty company name: use email domain
            if not company_name or company_name.lower() in ['nan', 'none']:
                company_name = email.split('@')[1].split('.')[0].capitalize()
            
            output_data.append({
                'company_name': company_name,
                'first_name': first_name,
                'email': email.lower()
            })
        
        if not output_data:
            print(f"  ⚠️ No valid rows found (all emails empty or invalid)")
            return 0
        
        # Create output DataFrame
        output_df = pd.DataFrame(output_data)
        
        # Remove duplicates by email
        original_count = len(output_df)
        output_df = output_df.drop_duplicates(subset=['email'], keep='first')
        if len(output_df) < original_count:
            print(f"  ⚠️ Removed {original_count - len(output_df)} duplicate emails")
        
        # Generate output filename
        output_filename = filename.replace('.xlsx', '.csv').replace('.xls', '.csv')
        output_path = os.path.join(output_dir, output_filename)
        
        # Save to CSV
        output_df.to_csv(output_path, index=False, encoding='utf-8')
        print(f"  ✅ Exported {len(output_df)} contacts to: {output_filename}")
        
        return len(output_df)
        
    except Exception as e:
        print(f"  ❌ Error processing {filename}: {e}")
        import traceback
        traceback.print_exc()
        return 0


def convert_folder(input_folder: str, output_folder: str = None) -> None:
    """
    Convert all Excel files in a folder to CSV format.
    
    Args:
        input_folder: Path to folder containing Excel files (e.g., 'cleaner' or '19-01-26')
        output_folder: Path to output folder (default: input_folder + '_output')
    """
    # Validate input folder
    if not os.path.exists(input_folder):
        print(f"❌ Input folder not found: {input_folder}")
        return
    
    # Set default output folder
    if output_folder is None:
        output_folder = input_folder + '_output'
    
    # Create output folder
    os.makedirs(output_folder, exist_ok=True)
    
    print(f"{'='*60}")
    print(f"Excel to CSV Converter")
    print(f"{'='*60}")
    print(f"📁 Input folder:  {input_folder}")
    print(f"📁 Output folder: {output_folder}")
    
    # Find all Excel files
    excel_files = [f for f in os.listdir(input_folder) 
                   if f.endswith(('.xlsx', '.xls')) and not f.startswith('~$')]
    
    if not excel_files:
        print(f"\n⚠️ No Excel files found in {input_folder}")
        return
    
    print(f"\n📊 Found {len(excel_files)} Excel file(s)")
    
    total_contacts = 0
    
    for excel_file in excel_files:
        excel_path = os.path.join(input_folder, excel_file)
        contacts = process_excel_file(excel_path, output_folder)
        total_contacts += contacts
    
    print(f"\n{'='*60}")
    print(f"✅ COMPLETE: Exported {total_contacts} total contacts")
    print(f"📁 CSV files saved to: {output_folder}")
    print(f"{'='*60}")


if __name__ == "__main__":
    import sys
    
    # Default folder name
    default_folder = "19-01-26"
    
    # Allow command line argument for folder name
    if len(sys.argv) > 1:
        input_folder = sys.argv[1]
    else:
        input_folder = default_folder
    
    # Run conversion
    convert_folder(input_folder)
