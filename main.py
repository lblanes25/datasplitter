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

def find_audit_leader_column(df):
    """Find the column that contains 'Audit Leader' in its name, 
    with preference for columns in the main data table rather than summary table."""
    # Check for exact column names first - these are likely in the main data table
    exact_matches = ['Audit Leader Name', 'Audit Leader Names', 'Audit Lead', 'Audit Leader']
    for col in df.columns:
        if col in exact_matches:
            return col
    
    # If no exact matches, check for columns that are both:
    # 1. Contain "audit leader" or "audit lead"
    # 2. Are in a row with other common data headers (suggesting it's the main table, not summary)
    audit_leader_cols = []
    
    for col in df.columns:
        if isinstance(col, str) and ('audit leader' in col.lower() or 'audit lead' in col.lower()):
            audit_leader_cols.append(col)
    
    # If we have multiple matches, try to identify the main data table
    if len(audit_leader_cols) > 1:
        # Look for columns that appear in rows with other common data headers
        common_data_headers = ['id', 'date', 'status', 'department', 'project', 'qa', 'finding', 
                              'risk', 'compliance', 'severity', 'due', 'assign']
        
        # Score each column based on how many common data headers appear in the same row
        col_scores = {}
        for col in audit_leader_cols:
            # Check how many other columns in this row contain common data headers
            other_cols = [c for c in df.columns if c != col]
            score = sum(1 for c in other_cols if any(header in str(c).lower() for header in common_data_headers))
            col_scores[col] = score
        
        # Return the column with the highest score (most likely to be in the main data table)
        if col_scores:
            best_col = max(col_scores.items(), key=lambda x: x[1])[0]
            return best_col
    
    # If only one match or can't determine from context, return the first match
    return audit_leader_cols[0] if audit_leader_cols else None

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
            # First, try to find the right header row by searching for patterns
            # that would indicate the main data table rather than summary tables
            found_main_data_table = False
            
            # Try reading different portions of the sheet to find the main data section
            for skip_rows in [0, 5, 10, 15, 20]:
                # Read a block of rows, assuming one might contain our header
                block_df = pd.read_excel(excel_file, sheet_name=sheet_name, 
                                        header=None, skiprows=skip_rows, nrows=10)
                
                # For each row in this block, check if it looks like a main data header row
                for i in range(min(10, len(block_df))):
                    row = block_df.iloc[i]
                    
                    # Count mentions of common data table headers to identify the main data table
                    common_headers = ['id', 'date', 'status', 'department', 'project', 'qa', 
                                     'finding', 'risk', 'compliance', 'severity', 'due']
                    
                    # Count how many cells in this row contain common header terms
                    header_matches = 0
                    audit_leader_idx = None
                    
                    for j, cell_value in enumerate(row):
                        if isinstance(cell_value, str):
                            cell_lower = cell_value.lower()
                            # Check for common headers
                            if any(header in cell_lower for header in common_headers):
                                header_matches += 1
                            
                            # Check for audit leader column
                            if 'audit' in cell_lower and ('leader' in cell_lower or 'lead' in cell_lower):
                                audit_leader_idx = j
                    
                    # If this row has both audit leader and several other common headers,
                    # it's likely the main data table header row
                    if audit_leader_idx is not None and header_matches >= 3:
                        print(f"  Found main data header row at skiprows={skip_rows}, row={i}, with {header_matches} common headers")
                        
                        # Read the sheet with this as the header row
                        df = pd.read_excel(excel_file, sheet_name=sheet_name, 
                                          skiprows=skip_rows+i, header=0)
                        
                        # Find the audit leader column
                        leader_col = find_audit_leader_column(df)
                        
                        if leader_col:
                            print(f"  Found main data leader column: {leader_col}")
                            leaders = df[leader_col].dropna().astype(str).unique()
                            # Filter out empty strings
                            leaders = [leader for leader in leaders if leader.strip()]
                            all_leaders.update(leaders)
                            print(f"  Added {len(leaders)} leaders from main data table")
                            found_main_data_table = True
                            break
                
                if found_main_data_table:
                    break
            
            # If we couldn't find the main data table with our heuristic approach,
            # fall back to the standard method
            if not found_main_data_table:
                print(f"  Could not identify main data table, trying standard approach")
                
                # Try reading with auto header detection
                df = pd.read_excel(excel_file, sheet_name=sheet_name)
                
                # Find the column with 'Audit Leader' in its name
                leader_col = find_audit_leader_column(df)
                
                # If found, add all non-null leaders to the set
                if leader_col:
                    print(f"  Found leader column: {leader_col} in sheet {sheet_name}")
                    leaders = df[leader_col].dropna().astype(str).unique()
                    # Filter out empty strings
                    leaders = [leader for leader in leaders if leader.strip()]
                    all_leaders.update(leaders)
                    print(f"  Added {len(leaders)} leaders from this sheet (standard method)")
        
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
                    # First, try to find the main data table (not summary tables)
                    main_data_found = False
                    
                    # Try different starting rows to skip past summary tables
                    for skiprows in [0, 5, 10, 15, 20]:
                        # Read a sample to look for header patterns
                        sample_df = pd.read_excel(input_path, sheet_name=sheet_name, 
                                                 header=None, skiprows=skiprows, nrows=10)
                        
                        for i in range(min(10, len(sample_df))):
                            row = sample_df.iloc[i]
                            
                            # Count how many common data headers we find in this row
                            common_headers = ['id', 'date', 'status', 'department', 'project', 'qa', 
                                             'finding', 'risk', 'compliance', 'severity', 'due']
                            header_count = 0
                            audit_leader_col = None
                            qa_results_col = None
                            
                            for j, cell_value in enumerate(row):
                                if isinstance(cell_value, str):
                                    cell_lower = cell_value.lower()
                                    # Check for common headers
                                    if any(header in cell_lower for header in common_headers):
                                        header_count += 1
                                    
                                    # Check for audit leader column
                                    if 'audit' in cell_lower and ('leader' in cell_lower or 'lead' in cell_lower):
                                        audit_leader_col = j
                                    
                                    # Check for QA Results column
                                    if 'qa' in cell_lower and 'result' in cell_lower:
                                        qa_results_col = j
                            
                            # If this row has both audit leader and several data headers,
                            # it's likely our main data table header
                            if audit_leader_col is not None and header_count >= 3:
                                print(f"    Found likely main data header at skiprows={skiprows}, row={i}")
                                
                                # Read the data with this as the header
                                header_row_index = skiprows + i
                                df = pd.read_excel(input_path, sheet_name=sheet_name, 
                                                  header=header_row_index)
                                
                                # Find the audit leader column
                                leader_col = find_audit_leader_column(df)
                                
                                if leader_col:
                                    print(f"    Using column: {leader_col}")
                                    
                                    # Filter for this leader
                                    filtered_df = filter_data_for_leader(df, leader_col, leader_name)
                                    
                                    if not filtered_df.empty:
                                        print(f"    Found {len(filtered_df)} rows for {leader_name}")
                                        found_data = True
                                        main_data_found = True
                                        
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
                                        for row_idx in range(1, header_row_index + 2):  # +2 for 1-indexing and to include header
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
                                        header_row_idx = header_row_index + 1  # +1 because Excel is 1-indexed
                                        for cell in new_sheet[header_row_idx]:
                                            cell.font = bold_font
                                        
                                        # Set column widths
                                        for i in range(1, len(row) + 1):
                                            new_sheet.column_dimensions[get_column_letter(i)].width = 15
                                        
                                        break  # Stop searching for headers in this block
                                
                            if main_data_found:
                                break
                        
                        if main_data_found:
                            break
                    
                    # If main data table wasn't found, try the standard approach
                    if not main_data_found:
                        print(f"    Trying standard approach (may include summary tables)")
                        
                        df = pd.read_excel(input_path, sheet_name=sheet_name)
                        leader_col = find_audit_leader_column(df)
                        
                        if leader_col:
                            print(f"    Found leader column: {leader_col}")
                            filtered_df = filter_data_for_leader(df, leader_col, leader_name)
                            
                            if not filtered_df.empty:
                                print(f"    Found {len(filtered_df)} rows for {leader_name} (standard method)")
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
                                
                                # Copy the header row
                                header_row = df.columns.tolist()
                                new_sheet.append(header_row)
                                
                                # Write data rows
                                for _, row in filtered_df.iterrows():
                                    new_sheet.append(row.tolist())
                                
                                # Apply formatting
                                from openpyxl.styles import Font
                                bold_font = Font(bold=True)
                                for cell in new_sheet[1]:  # Bold the header row
                                    cell.font = bold_font
                                    
                                # Set column widths
                                for i, _ in enumerate(header_row, 1):
                                    new_sheet.column_dimensions[get_column_letter(i)].width = 15
                            else:
                                print(f"    No data found for {leader_name} in sheet {sheet_name}")
                                
                                # Copy the first 15 rows (likely to include summary tables and headers)
                                original_sheet = original_wb[sheet_name]
                                rows_to_copy = min(15, len(list(original_sheet.rows)))
                                
                                for row_idx in range(1, rows_to_copy + 1):
                                    row_data = []
                                    for cell in original_sheet[row_idx]:
                                        row_data.append(cell.value)
                                    new_sheet.append(row_data)
                    
                    # If no specific approach worked, copy the structure of the sheet
                    if not found_data and not main_data_found:
                        print(f"    Unable to find data for {leader_name}, copying sheet structure")
                        
                        # Copy the first part of the sheet to maintain summary tables and headers
                        original_sheet = original_wb[sheet_name]
                        
                        # Determine how much to copy - either the whole sheet if small, or just the top portion
                        max_rows_to_copy = 20  # Copy up to this many rows to ensure headers are included
                        rows_to_copy = min(max_rows_to_copy, len(list(original_sheet.rows)))
                        
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
