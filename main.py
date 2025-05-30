import pandas as pd
import numpy as np
import os
import sys
import traceback
from openpyxl import load_workbook, Workbook
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.utils import get_column_letter

def is_qa_id_sheet(sheet_name):
    """Check if the sheet name follows the QA-ID-XXX format."""
    if not isinstance(sheet_name, str):
        return False
    return sheet_name.upper().startswith('QA-ID-')

def find_data_table_header_row(workbook_path, sheet_name):
    """Find the header row by looking for 'Detailed Results' in column B, 
    then finding 'Audit Leader' after it."""
    try:
        # Read the sheet without any header assumptions
        df_no_header = pd.read_excel(workbook_path, sheet_name=sheet_name, header=None)
        
        # Column B is index 1 (0-based)
        col_b = df_no_header.iloc[:, 1] if len(df_no_header.columns) > 1 else pd.Series()
        
        detailed_results_row = None
        audit_leader_row = None
        
        # First, find "Detailed Results" in column B
        for idx, value in col_b.items():
            if pd.notna(value) and isinstance(value, str):
                if 'detailed results' in value.lower():
                    detailed_results_row = idx
                    print(f"    Found 'Detailed Results' at row {idx + 1}")
                    break
        
        if detailed_results_row is None:
            print(f"    Could not find 'Detailed Results' in column B")
            return None
        
        # Now find "Audit Leader" after "Detailed Results"
        for idx in range(detailed_results_row + 1, len(col_b)):
            value = col_b.iloc[idx] if idx < len(col_b) else None
            if pd.notna(value) and isinstance(value, str):
                if value.strip().lower() == 'audit leader':
                    audit_leader_row = idx
                    print(f"    Found 'Audit Leader' header at row {idx + 1}")
                    break
        
        if audit_leader_row is None:
            print(f"    Could not find 'Audit Leader' after 'Detailed Results'")
            return None
            
        return audit_leader_row
        
    except Exception as e:
        print(f"    Error finding header row: {e}")
        return None

def find_audit_leader_column(df):
    """Find the column that contains 'Audit Leader' - should be straightforward now."""
    for col in df.columns:
        if isinstance(col, str) and col.strip().lower() == 'audit leader':
            return col
    
    # Fallback to partial match
    for col in df.columns:
        if isinstance(col, str) and 'audit leader' in col.lower():
            return col
            
    return None

def get_all_audit_leaders(workbook_path):
    """Collect all unique Audit Leaders from QA-ID-XXX sheets only."""
    print(f"Opening workbook to get audit leaders: {workbook_path}")
    
    all_leaders = set()
    
    # Using pandas to read each sheet
    excel_file = pd.ExcelFile(workbook_path)
    
    # Process only QA-ID-XXX sheets
    qa_sheets = [sheet for sheet in excel_file.sheet_names if is_qa_id_sheet(sheet)]
    print(f"Found {len(qa_sheets)} QA-ID sheets to process: {qa_sheets}")
    
    for sheet_name in qa_sheets:
        print(f"  Scanning sheet: {sheet_name} for audit leaders")
        
        try:
            # Find the header row using our new method
            header_row_idx = find_data_table_header_row(workbook_path, sheet_name)
            
            if header_row_idx is not None:
                # Read the data with the identified header row
                df = pd.read_excel(excel_file, sheet_name=sheet_name, 
                                  skiprows=header_row_idx, header=0)
                
                # Find the audit leader column
                leader_col = find_audit_leader_column(df)
                
                if leader_col:
                    print(f"  Found leader column: {leader_col}")
                    leaders = df[leader_col].dropna().astype(str).unique()
                    # Filter out empty strings and header text
                    leaders = [leader for leader in leaders if leader.strip() and leader.strip().lower() != 'audit leader']
                    all_leaders.update(leaders)
                    print(f"  Added {len(leaders)} leaders: {', '.join(leaders)}")
                else:
                    print(f"  Could not find audit leader column in sheet {sheet_name}")
            else:
                print(f"  Could not find data table in sheet {sheet_name}")
        
        except Exception as e:
            print(f"  Error processing sheet {sheet_name}: {e}")
            traceback.print_exc()
    
    # Convert to sorted list and print for verification
    leader_list = sorted(list(all_leaders))
    print(f"Total unique audit leaders found: {len(leader_list)}")
    print(f"Leaders: {', '.join(leader_list)}")
    return leader_list

def normalize_leader_name(name):
    """Normalize leader name for comparison by removing extra spaces and making lowercase."""
    if not isinstance(name, str):
        return str(name).lower().strip() if name is not None else ""
    return name.lower().strip()

def filter_data_for_leader(df, leader_col, leader_name):
    """Filter the dataframe to keep only rows for the specified audit leader."""
    if leader_col not in df.columns:
        return pd.DataFrame()  # Return empty dataframe if leader column not found
    
    normalized_leader = normalize_leader_name(leader_name)
    # Convert column to string, normalize, and compare
    normalized_col = df[leader_col].fillna("").astype(str).apply(normalize_leader_name)
    
    # Try exact match first
    filtered = df[normalized_col == normalized_leader].copy()
    
    # If no exact matches, try substring match
    if filtered.empty:
        filtered = df[normalized_col.str.contains(normalized_leader, na=False)].copy()
        
    return filtered

def copy_entire_sheet(source_sheet, target_sheet):
    """Copy an entire sheet from source to target, including formatting."""
    try:
        print(f"    Copying entire sheet: {source_sheet.title}")
        # Copy all rows and cell values
        for row in source_sheet.iter_rows():
            row_data = []
            for cell in row:
                row_data.append(cell.value)
            target_sheet.append(row_data)
        
        # Try to copy column widths
        for column in source_sheet.column_dimensions:
            if column in source_sheet.column_dimensions:
                target_sheet.column_dimensions[column].width = source_sheet.column_dimensions[column].width
        
        print(f"    Sheet copied successfully")
    except Exception as e:
        print(f"    Error copying sheet: {e}")
        traceback.print_exc()

def create_leader_workbook(input_path, output_path, leader_name):
    """Create a workbook with only data for the specified leader from QA-ID sheets, 
    copying non-QA-ID sheets entirely. Also highlight tabs red if they contain DNC results 
    for this leader, otherwise green."""
    print(f"\nProcessing workbook for: {leader_name}")
    
    try:
        # Load the original workbook to copy structure and styles
        print(f"  Loading original workbook: {input_path}")
        original_wb = load_workbook(input_path)
        new_wb = Workbook()
        
        # Remove the default sheet
        if "Sheet" in new_wb.sheetnames:
            del new_wb["Sheet"]
        
        # Track if we found any data for this leader
        found_data = False
        
        # Keep track of which sheets have DNC for tab coloring
        sheets_with_dnc = set()
        
        # Process each sheet
        for sheet_name in original_wb.sheetnames:
            print(f"  Processing sheet: {sheet_name}")
            
            # Create a new sheet with the same name for each sheet in original workbook
            new_sheet = new_wb.create_sheet(title=sheet_name)
            
                # Check if this is a QA-ID sheet that we need to filter
                if is_qa_id_sheet(sheet_name):
                    print(f"    This is a QA-ID sheet - filtering for {leader_name}")
                    
                    # Track if this sheet has DNC results for this leader
                    has_dnc = False
                    
                    # Try to find and filter data for this leader
                    try:
                        # Find the header row using our new method
                        header_row_idx = find_data_table_header_row(input_path, sheet_name)
                        
                        if header_row_idx is not None:
                            print(f"    Found data table header at row {header_row_idx + 1}")
                            
                            # Read the data with the identified header row
                            df = pd.read_excel(input_path, sheet_name=sheet_name, 
                                              skiprows=header_row_idx, header=0)
                            
                            # Find the audit leader column
                            leader_col = find_audit_leader_column(df)
                            
                            if leader_col:
                                print(f"    Using column: {leader_col}")
                                
                                # Filter for this leader
                                filtered_df = filter_data_for_leader(df, leader_col, leader_name)
                                
                                if not filtered_df.empty:
                                    print(f"    Found {len(filtered_df)} rows for {leader_name}")
                                    found_data = True
                                    
                                    # Check for QA Results column in the filtered data
                                    qa_col = None
                                    for col in filtered_df.columns:
                                        if isinstance(col, str) and 'qa' in col.lower() and 'result' in col.lower():
                                            qa_col = col
                                            break
                                    
                                    # If we found a QA Results column, check for DNC values
                                    if qa_col:
                                        print(f"    Found QA Results column: {qa_col}")
                                        # Check if any rows have "DNC" value for this leader
                                        dnc_rows = filtered_df[filtered_df[qa_col].astype(str).str.lower() == 'dnc']
                                        if not dnc_rows.empty:
                                            print(f"    ⚠️ Found {len(dnc_rows)} DNC results in this sheet!")
                                            has_dnc = True
                                            sheets_with_dnc.add(sheet_name)
                                    
                                    # Copy everything up to the header row from original
                                    # This preserves any summary tables at the top
                                    original_sheet = original_wb[sheet_name]
                                    for row_idx in range(1, header_row_idx + 2):  # +2 for 1-indexing and to include header
                                        row_data = []
                                        for cell in original_sheet[row_idx]:
                                            row_data.append(cell.value)
                                        new_sheet.append(row_data)
                                    
                                    # Add the filtered data rows 
                                    for _, row in filtered_df.iterrows():
                                        new_sheet.append(row.tolist())
                                    
                                    # Apply formatting
                                    from openpyxl.styles import Font
                                    bold_font = Font(bold=True)
                                    
                                    # Bold the header row
                                    header_row_idx_excel = header_row_idx + 1  # +1 because Excel is 1-indexed
                                    for cell in new_sheet[header_row_idx_excel]:
                                        cell.font = bold_font
                                    
                                    # Set column widths
                                    for i in range(1, len(row) + 1):
                                        new_sheet.column_dimensions[get_column_letter(i)].width = 15
                                    
                                else:
                                    print(f"    No data found for {leader_name} in sheet {sheet_name}")
                                    
                                    # Copy the first part including headers
                                    original_sheet = original_wb[sheet_name]
                                    rows_to_copy = min(header_row_idx + 10, len(list(original_sheet.rows)))
                                    
                                    for row_idx in range(1, rows_to_copy + 1):
                                        row_data = []
                                        for cell in original_sheet[row_idx]:
                                            row_data.append(cell.value)
                                        new_sheet.append(row_data)
                            else:
                                print(f"    Could not find audit leader column")
                                # Copy structure
                                original_sheet = original_wb[sheet_name]
                                rows_to_copy = min(20, len(list(original_sheet.rows)))
                                
                                for row_idx in range(1, rows_to_copy + 1):
                                    row_data = []
                                    for cell in original_sheet[row_idx]:
                                        row_data.append(cell.value)
                                    new_sheet.append(row_data)
                        else:
                            print(f"    Could not find data table header")
                            # Copy structure
                            original_sheet = original_wb[sheet_name]
                            rows_to_copy = min(20, len(list(original_sheet.rows)))
                            
                            for row_idx in range(1, rows_to_copy + 1):
                                row_data = []
                                for cell in original_sheet[row_idx]:
                                    row_data.append(cell.value)
                                new_sheet.append(row_data)
                        
                    except Exception as e:
                        print(f"    Error processing QA-ID sheet {sheet_name}: {e}")
                        traceback.print_exc()
                        
                        # On error, copy at least part of the sheet to maintain structure
                        original_sheet = original_wb[sheet_name]
                        rows_to_copy = min(15, len(list(original_sheet.rows)))
                        
                        for row_idx in range(1, rows_to_copy + 1):
                            try:
                                row_data = []
                                for cell in original_sheet[row_idx]:
                                    row_data.append(cell.value)
                                new_sheet.append(row_data)
                            except:
                                break  # If we can't copy this row, stop trying
            
            else:
                # This is NOT a QA-ID sheet - copy it entirely as-is
                print(f"    This is not a QA-ID sheet - copying entirely")
                try:
                    original_sheet = original_wb[sheet_name]
                    copy_entire_sheet(original_sheet, new_sheet)
                except Exception as e:
                    print(f"    Error copying non-QA-ID sheet {sheet_name}: {e}")
                    traceback.print_exc()
        
        # Color the tabs based on DNC results
        from openpyxl.styles import PatternFill
        
        for sheet_name in new_wb.sheetnames:
            # Get the sheet
            sheet = new_wb[sheet_name]
            
            # Only color QA-ID sheets based on DNC status
            if is_qa_id_sheet(sheet_name):
                # Set tab color
                if sheet_name in sheets_with_dnc:
                    # Red tab for sheets with DNC
                    sheet.sheet_properties.tabColor = "FF0000"  # Red
                    print(f"  🔴 Tab for QA-ID sheet {sheet_name} colored RED (has DNC)")
                else:
                    # Green tab for sheets without DNC
                    sheet.sheet_properties.tabColor = "00FF00"  # Green
                    print(f"  🟢 Tab for QA-ID sheet {sheet_name} colored GREEN (no DNC)")
            else:
                # Leave non-QA-ID sheets with default tab color
                print(f"  ⚪ Tab for non-QA-ID sheet {sheet_name} left with default color")
        
        # Save the workbook
        print(f"  Saving workbook to: {output_path}")
        new_wb.save(output_path)
        return found_data
    except Exception as e:
        print(f"Error creating workbook for {leader_name}: {e}")
        traceback.print_exc()
        return False

def split_audit_data(input_file, output_dir='audit_leaders'):
    """Split the audit data file into separate files for each audit leader."""
    try:
        # Create the output directory if it doesn't exist
        if not os.path.exists(output_dir):
            print(f"Creating output directory: {output_dir}")
            os.makedirs(output_dir)
        else:
            print(f"Output directory already exists: {output_dir}")
        
        # Get all audit leaders from QA-ID sheets only
        audit_leaders = get_all_audit_leaders(input_file)
        print(f"Found {len(audit_leaders)} audit leaders across QA-ID sheets")
        
        # Create a workbook for each leader
        success_count = 0
        for leader in audit_leaders:
            # Create filename as requested: "Audit Leader + Horizontal Review 2025"
            output_file = os.path.join(output_dir, f"{leader} Horizontal Review 2025.xlsx")
            
            print(f"Creating file for {leader}...")
            result = create_leader_workbook(input_file, output_file, leader)
            
            if result:
                success_count += 1
                print(f"  Saved to {output_file}")
            else:
                print(f"  No data found for {leader}, but workbook created with structure")
        
        print(f"\nCompleted: Created {success_count} files for {len(audit_leaders)} audit leaders!")
        return True
    except Exception as e:
        print(f"Error in split_audit_data: {e}")
        traceback.print_exc()
        return False

if __name__ == "__main__":
    print("Script starting...")
    try:
        # Command line arguments parsing
        import argparse
        
        parser = argparse.ArgumentParser(description="Split audit data Excel file by audit leaders.")
        parser.add_argument("-i", "--input", help="Path to input Excel file")
        parser.add_argument("-o", "--output", help="Path to output directory")
        args = parser.parse_args()
        
        # Get input file - from args, or prompt user
        input_file = args.input
        if not input_file:
            input_file = input("Enter path to Excel file: ").strip()
            # Remove quotes if user copied path with quotes
            input_file = input_file.strip('"\'')
        
        # Check if input file exists
        print(f"Checking for input file: {input_file}")
        if not os.path.exists(input_file):
            print(f"ERROR: Input file not found at {input_file}")
            sys.exit(1)
        
        # Get output directory - from args, prompt user, or default
        output_dir = args.output
        if not output_dir:
            output_dir = input("Enter output directory path (press Enter for default 'audit_leaders'): ").strip()
            # Remove quotes if user copied path with quotes
            output_dir = output_dir.strip('"\'')
            if not output_dir:
                # Default output directory - in same location as input file
                input_dir = os.path.dirname(input_file)
                output_dir = os.path.join(input_dir, "audit_leaders")
                print(f"Using default output directory: {output_dir}")
        
        # Run the script
        print(f"Input file found. Starting to process...")
        if split_audit_data(input_file, output_dir):
            print("Script completed successfully!")
        else:
            print("Script completed with errors!")
    except Exception as e:
        print(f"ERROR: {str(e)}")
        traceback.print_exc()
