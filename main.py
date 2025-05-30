import os
import shutil
import re
from pathlib import Path
import pandas as pd
import openpyxl
from openpyxl.styles import PatternFill
from typing import Dict, List, Tuple, Optional, Set
import logging

# Set up logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

def sanitize_filename(name: str) -> str:
    """Clean a string to make it safe for use as a Windows filename."""
    # Replace invalid characters with underscores
    return re.sub(r'[<>:"/\\|?*\n\r]+', '_', name).strip()

def find_table_boundaries(sheet, sheet_name: str) -> Optional[Tuple[int, int, int, int]]:
    """
    Find the data table boundaries in a QA-ID sheet.
    
    Returns:
        Tuple of (header_row, data_start_row, data_end_row, max_col) or None if not found
    """
    try:
        # Step 1: Find "Detailed Results" in column B
        detailed_results_row = None
        for row_num in range(1, sheet.max_row + 1):
            cell_value = sheet.cell(row=row_num, column=2).value
            if cell_value and "Detailed Results" in str(cell_value):
                detailed_results_row = row_num
                logger.info(f"Found 'Detailed Results' at row {row_num} in {sheet_name}")
                break
        
        if detailed_results_row is None:
            logger.warning(f"Could not find 'Detailed Results' in column B for sheet {sheet_name}")
            return None
        
        # Step 2: Find "Audit Leader" in column B after "Detailed Results"
        audit_leader_row = None
        for row_num in range(detailed_results_row + 1, sheet.max_row + 1):
            cell_value = sheet.cell(row=row_num, column=2).value
            if cell_value and "Audit Leader" in str(cell_value):
                audit_leader_row = row_num
                logger.info(f"Found 'Audit Leader' at row {row_num} in {sheet_name}")
                break
        
        if audit_leader_row is None:
            logger.warning(f"Could not find 'Audit Leader' in column B after 'Detailed Results' for sheet {sheet_name}")
            return None
        
        # Step 3: Determine data boundaries
        header_row = audit_leader_row
        data_start_row = header_row + 1
        
        # Find the end of data by looking for first completely empty row
        data_end_row = sheet.max_row
        for row_num in range(data_start_row, sheet.max_row + 1):
            row_has_data = False
            for col_num in range(1, sheet.max_column + 1):
                if sheet.cell(row=row_num, column=col_num).value is not None:
                    row_has_data = True
                    break
            if not row_has_data:
                data_end_row = row_num - 1
                break
        
        # Find the actual max column used in the header row
        max_col = 1
        for col_num in range(1, sheet.max_column + 1):
            if sheet.cell(row=header_row, column=col_num).value is not None:
                max_col = col_num
        
        logger.info(f"Table boundaries for {sheet_name}: header_row={header_row}, data_start={data_start_row}, data_end={data_end_row}, max_col={max_col}")
        return header_row, data_start_row, data_end_row, max_col
        
    except Exception as e:
        logger.error(f"Error finding table boundaries in {sheet_name}: {str(e)}")
        return None

def get_column_mapping(sheet, header_row: int, max_col: int) -> Dict[str, int]:
    """
    Create a mapping of column names to column numbers from the header row.
    """
    column_mapping = {}
    for col_num in range(1, max_col + 1):
        header_value = sheet.cell(row=header_row, column=col_num).value
        if header_value:
            column_mapping[str(header_value).strip()] = col_num
    
    logger.info(f"Column mapping: {column_mapping}")
    return column_mapping

def extract_data_to_dataframe(sheet, header_row: int, data_start_row: int, data_end_row: int, max_col: int) -> pd.DataFrame:
    """
    Extract data from the sheet into a pandas DataFrame.
    """
    # Get headers
    headers = []
    for col_num in range(1, max_col + 1):
        header_value = sheet.cell(row=header_row, column=col_num).value
        headers.append(str(header_value) if header_value is not None else f"Column_{col_num}")
    
    # Get data
    data = []
    for row_num in range(data_start_row, data_end_row + 1):
        row_data = []
        for col_num in range(1, max_col + 1):
            cell_value = sheet.cell(row=row_num, column=col_num).value
            row_data.append(cell_value)
        data.append(row_data)
    
    df = pd.DataFrame(data, columns=headers)
    logger.info(f"Extracted DataFrame with shape {df.shape}")
    return df

def finalize_sheet_presentation(sheet):
    """
    Set sheet to A1 position and collapse all grouped rows/columns before saving.
    """
    try:
        # Set active cell to A1
        sheet.active_cell = "A1"
        
        # Set sheet view to show A1 (top-left corner)
        sheet.sheet_view.topLeftCell = "A1"
        
        # Collapse all row groups (outline levels)
        if hasattr(sheet, 'row_dimensions'):
            for row_num, row_dim in sheet.row_dimensions.items():
                if hasattr(row_dim, 'outline_level') and row_dim.outline_level > 0:
                    row_dim.hidden = True
        
        # Collapse all column groups (outline levels)  
        if hasattr(sheet, 'column_dimensions'):
            for col_letter, col_dim in sheet.column_dimensions.items():
                if hasattr(col_dim, 'outline_level') and col_dim.outline_level > 0:
                    col_dim.hidden = True
        
        # Alternative approach: Set outline summary below/right to collapse groups
        if hasattr(sheet, 'sheet_properties'):
            if hasattr(sheet.sheet_properties, 'outline_pr'):
                sheet.sheet_properties.outline_pr.summary_below = False
                sheet.sheet_properties.outline_pr.summary_right = False
        
        logger.info(f"Finalized presentation for sheet: {sheet.title}")
        
    except Exception as e:
        logger.warning(f"Could not finalize presentation for sheet {sheet.title}: {str(e)}")
        # Don't fail the whole process if presentation cleanup fails
def write_dataframe_to_sheet(sheet, df: pd.DataFrame, data_start_row: int, data_end_row: int, max_col: int):
    """
    Replace existing data rows with filtered DataFrame data by deleting rows and inserting new ones.
    """
    # Step 1: Delete all existing data rows (from bottom to top to avoid index shifting)
    rows_to_delete = data_end_row - data_start_row + 1
    for i in range(rows_to_delete):
        sheet.delete_rows(data_start_row)
    
    # Step 2: Insert new rows for the filtered data
    for idx, row in df.iterrows():
        excel_row = data_start_row + idx
        sheet.insert_rows(excel_row)
        
        # Write the data to the new row
        for col_idx, value in enumerate(row):
            if col_idx < max_col:  # Don't exceed original column range
                sheet.cell(row=excel_row, column=col_idx + 1).value = value

def filter_and_sort_data(df: pd.DataFrame, audit_leader: str, column_mapping: Dict[str, int]) -> Tuple[pd.DataFrame, bool]:
    """
    Filter data for specific audit leader and sort with DNC values first.
    
    Returns:
        Tuple of (filtered_df, has_dnc_values)
    """
    # Filter for the specific audit leader
    if "Audit Leader" not in column_mapping:
        logger.error("'Audit Leader' column not found in column mapping")
        return df, False
    
    audit_leader_col = None
    for col_name in df.columns:
        if "Audit Leader" in str(col_name):
            audit_leader_col = col_name
            break
    
    if audit_leader_col is None:
        logger.error("Could not find Audit Leader column in DataFrame")
        return df, False
    
    # Filter rows for this audit leader
    filtered_df = df[df[audit_leader_col] == audit_leader].copy()
    logger.info(f"Filtered to {len(filtered_df)} rows for audit leader: {audit_leader}")
    
    if len(filtered_df) == 0:
        return filtered_df, False
    
    # Look for the specific Overall Test Result column
    result_col_name = "Overall Test Result (after considering any applicable test result overrides)"
    
    if result_col_name not in filtered_df.columns:
        logger.warning(f"Could not find '{result_col_name}' column")
        # Try to find a close match in case of slight variations
        for col_name in filtered_df.columns:
            if "Overall Test Result" in str(col_name):
                result_col_name = col_name
                logger.info(f"Using column '{result_col_name}' instead")
                break
        else:
            logger.error("No Overall Test Result column found - cannot check for DNC values")
            return filtered_df, False
    
    # Check for DNC values in the specific column
    has_dnc = False
    if result_col_name in filtered_df.columns:
        dnc_count = filtered_df[result_col_name].astype(str).str.contains("DNC", case=False, na=False).sum()
        if dnc_count > 0:
            has_dnc = True
            logger.info(f"Found {dnc_count} DNC values in '{result_col_name}' column")
    
    # Sort with DNC values first
    if has_dnc:
        # Create a sort key that prioritizes DNC values
        def sort_key(row):
            if result_col_name in row and "DNC" in str(row[result_col_name]):
                return 0  # DNC values get priority
            return 1  # Non-DNC values come after
        
        filtered_df['_sort_key'] = filtered_df.apply(sort_key, axis=1)
        filtered_df = filtered_df.sort_values('_sort_key').drop('_sort_key', axis=1)
        logger.info("Sorted data with DNC values first")
    
    return filtered_df, has_dnc

def analyze_workbook_structure(workbook_path: str) -> Tuple[Set[str], Dict[str, Tuple]]:
    """
    Analyze workbook structure once to get audit leaders and table boundaries.
    
    Returns:
        Tuple of (audit_leaders_set, sheet_info_dict)
        where sheet_info_dict maps sheet_name -> (header_row, data_start_row, data_end_row, max_col, column_mapping)
    """
    audit_leaders = set()
    sheet_info = {}
    
    try:
        wb = openpyxl.load_workbook(workbook_path, data_only=False)
        
        for sheet_name in wb.sheetnames:
            if sheet_name.startswith("QA-ID-"):
                logger.info(f"Analyzing sheet {sheet_name}")
                sheet = wb[sheet_name]
                
                # Find table boundaries
                boundaries = find_table_boundaries(sheet, sheet_name)
                if boundaries is None:
                    continue
                
                header_row, data_start_row, data_end_row, max_col = boundaries
                column_mapping = get_column_mapping(sheet, header_row, max_col)
                
                # Store sheet info for reuse
                sheet_info[sheet_name] = (header_row, data_start_row, data_end_row, max_col, column_mapping)
                
                # Find Audit Leader column
                audit_leader_col = None
                for col_name, col_num in column_mapping.items():
                    if "Audit Leader" in col_name:
                        audit_leader_col = col_num
                        break
                
                if audit_leader_col is None:
                    logger.warning(f"Could not find Audit Leader column in {sheet_name}")
                    continue
                
                # Extract unique audit leaders from this sheet
                for row_num in range(data_start_row, data_end_row + 1):
                    leader_value = sheet.cell(row=row_num, column=audit_leader_col).value
                    if leader_value and str(leader_value).strip():
                        audit_leaders.add(str(leader_value).strip())
        
        wb.close()
        
    except Exception as e:
        logger.error(f"Error analyzing workbook structure: {str(e)}")
    
    logger.info(f"Found unique audit leaders: {sorted(audit_leaders)}")
    logger.info(f"Analyzed {len(sheet_info)} QA-ID sheets")
    return audit_leaders, sheet_info

def process_workbook_by_audit_leaders(source_file: str, output_dir: str = None) -> Dict[str, str]:
    """
    Process an Excel workbook to create filtered versions for each audit leader.
    
    Args:
        source_file: Path to the source Excel file
        output_dir: Directory to save output files (defaults to same directory as source)
    
    Returns:
        Dictionary mapping audit leader names to output file paths
    """
    source_path = Path(source_file)
    if not source_path.exists():
        raise FileNotFoundError(f"Source file not found: {source_file}")
    
    if output_dir is None:
        output_dir = source_path.parent
    else:
        output_dir = Path(output_dir)
        output_dir.mkdir(parents=True, exist_ok=True)
    
    # Get the base filename without extension
    base_name = source_path.stem
    
    # OPTIMIZATION: Analyze workbook structure once
    logger.info("Analyzing workbook structure...")
    audit_leaders, sheet_info = analyze_workbook_structure(source_file)
    
    if not audit_leaders:
        logger.warning("No audit leaders found in the workbook")
        return {}
    
    if not sheet_info:
        logger.warning("No QA-ID sheets with valid table structure found")
        return {}
    
    results = {}
    
    # Process each audit leader using pre-analyzed structure
    for audit_leader in sorted(audit_leaders):
        logger.info(f"Processing workbook for audit leader: {audit_leader}")
        
        try:
            # Step 1: Copy the original file
            safe_leader_name = sanitize_filename(audit_leader)
            output_filename = f"{base_name} - {safe_leader_name}.xlsx"
            output_path = output_dir / output_filename
            shutil.copyfile(source_file, output_path)
            logger.info(f"Created copy: {output_path}")
            
            # Step 2: Open the copied file
            wb = openpyxl.load_workbook(output_path, data_only=False)
            
            # Step 3: Process each QA-ID sheet using pre-analyzed info
            for sheet_name, (header_row, data_start_row, data_end_row, max_col, column_mapping) in sheet_info.items():
                logger.info(f"Processing sheet: {sheet_name}")
                sheet = wb[sheet_name]
                
                # Extract data to DataFrame using known boundaries
                df = extract_data_to_dataframe(sheet, header_row, data_start_row, data_end_row, max_col)
                
                # Filter and sort data
                filtered_df, has_dnc = filter_and_sort_data(df, audit_leader, column_mapping)
                
                # Write filtered data back to sheet
                write_dataframe_to_sheet(sheet, filtered_df, data_start_row, data_end_row, max_col)
                
                # Set tab color based on DNC presence
                if has_dnc:
                    sheet.sheet_properties.tabColor = "FF0000"  # Red
                    logger.info(f"Set {sheet_name} tab color to red (DNC values present)")
                else:
                    sheet.sheet_properties.tabColor = "00FF00"  # Green
                    logger.info(f"Set {sheet_name} tab color to green (no DNC values)")
            
            # Step 4: Finalize all sheets for better presentation
            for sheet_name in wb.sheetnames:
                sheet = wb[sheet_name]
                finalize_sheet_presentation(sheet)
            
            # Step 5: Save the processed workbook
            wb.save(output_path)
            wb.close()
            
            results[audit_leader] = str(output_path)
            logger.info(f"Successfully processed workbook for {audit_leader}")
            
        except Exception as e:
            logger.error(f"Error processing workbook for {audit_leader}: {str(e)}")
            # Clean up partial file if it exists
            if output_path.exists():
                output_path.unlink()
    
    logger.info(f"Processing complete. Created {len(results)} workbooks.")
    return results

# Example usage
if __name__ == "__main__":
    # Example usage
    source_file = "your_workbook.xlsx"
    output_directory = "output_files"
    
    try:
        results = process_workbook_by_audit_leaders(source_file, output_directory)
        
        print("Processing Results:")
        for audit_leader, output_file in results.items():
            print(f"  {audit_leader}: {output_file}")
            
    except Exception as e:
        print(f"Error: {e}")
