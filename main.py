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

def get_result_column_number(sheet, header_row: int, max_col: int) -> Optional[int]:
    """Find the result column number using the same logic as before."""
    target_text = "overall test result (after considering any applicable test result overrides)"
    
    for col_num in range(1, max_col + 1):
        header_value = sheet.cell(row=header_row, column=col_num).value
        if header_value:
            normalized_col = ' '.join(str(header_value).replace('\n', ' ').split()).lower()
            if normalized_col == target_text:
                return col_num
    
    # Fallback logic
    for col_num in range(1, max_col + 1):
        header_value = sheet.cell(row=header_row, column=col_num).value
        if header_value:
            normalized_col = ' '.join(str(header_value).replace('\n', ' ').split()).lower()
            if ("overall test result" in normalized_col and 
                "considering" in normalized_col and 
                "applicable" in normalized_col and
                "override" not in normalized_col):
                return col_num
    
    return None

def get_audit_leader_column_number(column_mapping: Dict[str, int]) -> int:
    """Find the column number for Audit Leader."""
    for col_name, col_num in column_mapping.items():
        if "Audit Leader" in col_name:
            return col_num
    raise ValueError("Audit Leader column not found in column mapping")

def sort_sheet_by_dnc(sheet, header_row: int, data_start_row: int, data_end_row: int, result_col_num: int):
    """
    Sort all data in the sheet to put DNC values first, regardless of audit leader.
    This is done once per sheet before any filtering.
    """
    if not result_col_num:
        logger.info("No result column found - skipping DNC sorting")
        return
    
    # Read all row data
    dnc_rows = []
    non_dnc_rows = []
    
    for row_num in range(data_start_row, data_end_row + 1):
        # Read entire row
        row_data = []
        for col_num in range(1, sheet.max_column + 1):
            cell_value = sheet.cell(row=row_num, column=col_num).value
            row_data.append(cell_value)
        
        # Check if this row has DNC in the result column
        result_value = sheet.cell(row=row_num, column=result_col_num).value
        if result_value and "DNC" in str(result_value).upper():
            dnc_rows.append(row_data)
        else:
            non_dnc_rows.append(row_data)
    
    # Combine: DNC rows first, then non-DNC rows
    sorted_data = dnc_rows + non_dnc_rows
    
    # Write sorted data back to sheet
    for idx, row_data in enumerate(sorted_data):
        excel_row = data_start_row + idx
        for col_idx, value in enumerate(row_data):
            sheet.cell(row=excel_row, column=col_idx + 1).value = value
    
    logger.info(f"Sorted sheet: {len(dnc_rows)} DNC rows moved to top, {len(non_dnc_rows)} non-DNC rows below")

def filter_rows_by_audit_leader(sheet, audit_leader: str, data_start_row: int, data_end_row: int, audit_leader_col_num: int, result_col_num: int = None) -> bool:
    """
    Simple filtering: delete rows that don't match the audit leader.
    Assumes data is already sorted with DNC values first.
    
    Returns:
        bool: True if any DNC values remain after filtering
    """
    rows_to_delete = []
    has_dnc_remaining = False
    
    # Identify rows to delete and check for remaining DNC values
    for row_num in range(data_start_row, data_end_row + 1):
        audit_leader_cell = sheet.cell(row=row_num, column=audit_leader_col_num).value
        
        if str(audit_leader_cell).strip() == audit_leader:
            # This row matches - check if it has DNC
            if result_col_num:
                result_cell = sheet.cell(row=row_num, column=result_col_num).value
                if result_cell and "DNC" in str(result_cell).upper():
                    has_dnc_remaining = True
        else:
            # This row doesn't match - mark for deletion
            rows_to_delete.append(row_num)
    
    # Delete non-matching rows (from bottom to top to avoid index shifting)
    for row_num in reversed(rows_to_delete):
        sheet.delete_rows(row_num)
    
    logger.info(f"Filtered sheet: deleted {len(rows_to_delete)} non-matching rows, DNC remaining: {has_dnc_remaining}")
    return has_dnc_remaining

def analyze_workbook_structure(workbook_path: str) -> Tuple[Set[str], Dict[str, Tuple]]:
    """
    Analyze workbook structure once to get audit leaders and table boundaries.
    
    Returns:
        Tuple of (audit_leaders_set, sheet_info_dict)
        where sheet_info_dict maps sheet_name -> (header_row, data_start_row, data_end_row, max_col, column_mapping, result_col_name)
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
                
                # Extract data to identify result column and audit leaders
                df = extract_data_to_dataframe(sheet, header_row, data_start_row, data_end_row, max_col)
                
                # Find the correct result column once
                result_col_name = None
                target_text = "overall test result (after considering any applicable test result overrides)"
                
                for col_name in df.columns:
                    # Normalize the column name: remove newlines, extra spaces, convert to lowercase
                    normalized_col = ' '.join(str(col_name).replace('\n', ' ').split()).lower()
                    
                    if normalized_col == target_text:
                        result_col_name = col_name
                        logger.info(f"Found target result column in {sheet_name}: '{col_name}'")
                        break
                
                # Fallback: look for column containing key phrases and NOT containing "override"
                if result_col_name is None:
                    for col_name in df.columns:
                        normalized_col = ' '.join(str(col_name).replace('\n', ' ').split()).lower()
                        if ("overall test result" in normalized_col and 
                            "considering" in normalized_col and 
                            "applicable" in normalized_col and
                            "override" not in normalized_col):
                            result_col_name = col_name
                            logger.info(f"Found fallback result column in {sheet_name}: '{col_name}'")
                            break
                
                if result_col_name is None:
                    logger.warning(f"Could not find result column in {sheet_name}")
                
                # Store sheet info including the result column for reuse
                sheet_info[sheet_name] = (header_row, data_start_row, data_end_row, max_col, column_mapping, result_col_name)
                
                # Find Audit Leader column and extract unique values
                audit_leader_col = None
                for col_name, col_num in column_mapping.items():
                    if "Audit Leader" in col_name:
                        audit_leader_col = col_name
                        break
                
                if audit_leader_col is None:
                    logger.warning(f"Could not find Audit Leader column in {sheet_name}")
                    continue
                
                # Extract unique audit leaders from this sheet
                for leader_value in df[audit_leader_col].dropna().unique():
                    if str(leader_value).strip():
                        audit_leaders.add(str(leader_value).strip())
        
        wb.close()
        
    except Exception as e:
        logger.error(f"Error analyzing workbook structure: {str(e)}")
    
    logger.info(f"Found unique audit leaders: {sorted(audit_leaders)}")
    logger.info(f"Analyzed {len(sheet_info)} QA-ID sheets")
    return audit_leaders, sheet_info

def pre_sort_workbook_by_dnc(source_file: str) -> str:
    """
    Create a pre-sorted version of the workbook with DNC values at the top of each sheet.
    This sorted file will be used as the base for all audit leader filtering.
    
    Returns:
        Path to the pre-sorted workbook
    """
    source_path = Path(source_file)
    sorted_filename = f"{source_path.stem}_sorted_dnc{source_path.suffix}"
    sorted_path = source_path.parent / sorted_filename
    
    # Copy original file
    shutil.copyfile(source_file, sorted_path)
    logger.info(f"Created pre-sort copy: {sorted_path}")
    
    # Analyze structure
    audit_leaders, sheet_info = analyze_workbook_structure(source_file)
    
    # Open the copied file and sort each sheet
    wb = openpyxl.load_workbook(sorted_path, data_only=False)
    
    for sheet_name, (header_row, data_start_row, data_end_row, max_col, column_mapping, result_col_name) in sheet_info.items():
        logger.info(f"Sorting sheet by DNC: {sheet_name}")
        sheet = wb[sheet_name]
        
        # Get result column number
        result_col_num = get_result_column_number(sheet, header_row, max_col)
        
        # Sort this sheet by DNC
        sort_sheet_by_dnc(sheet, header_row, data_start_row, data_end_row, result_col_num)
    
    # Save the pre-sorted workbook
    wb.save(sorted_path)
    wb.close()
    logger.info(f"Pre-sorted workbook saved: {sorted_path}")
    
    return str(sorted_path)

def process_workbook_by_audit_leaders(source_file: str, output_dir: str = None) -> Dict[str, str]:
    """
    Process an Excel workbook to create filtered versions for each audit leader.
    Optimized approach: Sort once by DNC, then filter many times by audit leader.
    
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
    
    base_name = source_path.stem
    
    # Step 1: Analyze workbook structure
    logger.info("Analyzing workbook structure...")
    audit_leaders, sheet_info = analyze_workbook_structure(source_file)
    
    if not audit_leaders:
        logger.warning("No audit leaders found in the workbook")
        return {}
    
    if not sheet_info:
        logger.warning("No QA-ID sheets with valid table structure found")
        return {}
    
    # Step 2: Create pre-sorted workbook (DNC values first)
    logger.info("Pre-sorting workbook by DNC values...")
    sorted_workbook_path = pre_sort_workbook_by_dnc(source_file)
    
    results = {}
    
    # Step 3: Process each audit leader using the pre-sorted workbook
    for audit_leader in sorted(audit_leaders):
        logger.info(f"Processing workbook for audit leader: {audit_leader}")
        
        try:
            # Copy the pre-sorted file (not the original)
            safe_leader_name = sanitize_filename(audit_leader)
            output_filename = f"{base_name} - {safe_leader_name}.xlsx"
            output_path = output_dir / output_filename
            shutil.copyfile(sorted_workbook_path, output_path)
            logger.info(f"Created copy from pre-sorted file: {output_path}")
            
            # Open the copied file
            wb = openpyxl.load_workbook(output_path, data_only=False)
            
            # Process each QA-ID sheet (simple filtering only)
            for sheet_name, (header_row, data_start_row, data_end_row, max_col, column_mapping, result_col_name) in sheet_info.items():
                logger.info(f"Filtering sheet: {sheet_name}")
                sheet = wb[sheet_name]
                
                # Get column numbers
                audit_leader_col_num = get_audit_leader_column_number(column_mapping)
                result_col_num = get_result_column_number(sheet, header_row, max_col)
                
                # Simple filter by audit leader (no sorting needed - already done)
                has_dnc = filter_rows_by_audit_leader(
                    sheet, audit_leader, data_start_row, data_end_row, 
                    audit_leader_col_num, result_col_num
                )
                
                # Set tab color based on DNC presence
                if has_dnc:
                    sheet.sheet_properties.tabColor = "FF0000"  # Red
                    logger.info(f"Set {sheet_name} tab color to red (DNC values present)")
                else:
                    sheet.sheet_properties.tabColor = "00FF00"  # Green
                    logger.info(f"Set {sheet_name} tab color to green (no DNC values)")
            
            # Finalize all sheets
            for sheet_name in wb.sheetnames:
                sheet = wb[sheet_name]
                finalize_sheet_presentation(sheet)
            
            # Save the processed workbook
            wb.save(output_path)
            wb.close()
            
            results[audit_leader] = str(output_path)
            logger.info(f"Successfully processed workbook for {audit_leader}")
            
        except Exception as e:
            logger.error(f"Error processing workbook for {audit_leader}: {str(e)}")
            if output_path.exists():
                output_path.unlink()
    
    # Step 4: Clean up the temporary pre-sorted file
    try:
        Path(sorted_workbook_path).unlink()
        logger.info("Cleaned up temporary pre-sorted file")
    except Exception as e:
        logger.warning(f"Could not clean up temporary file {sorted_workbook_path}: {str(e)}")
    
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
