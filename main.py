import pandas as pd
import numpy as np
import os
import sys
import traceback
from openpyxl import load_workbook, Workbook
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.utils import get_column_letter
from openpyxl.styles import PatternFill, Font


def is_qa_id_sheet(sheet_name):
    if not isinstance(sheet_name, str):
        return False
    return sheet_name.upper().startswith('QA-ID-')


def find_data_table_header_row(workbook_path, sheet_name):
    try:
        df_no_header = pd.read_excel(workbook_path, sheet_name=sheet_name, header=None)
        col_b = df_no_header.iloc[:, 1] if len(df_no_header.columns) > 1 else pd.Series()

        detailed_results_row = None
        audit_leader_row = None

        for idx, value in col_b.items():
            if pd.notna(value) and isinstance(value, str):
                if 'detailed results' in value.lower():
                    detailed_results_row = idx
                    print(f"    Found 'Detailed Results' at row {idx + 1}")
                    break

        if detailed_results_row is None:
            print(f"    Could not find 'Detailed Results' in column B")
            return None

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
    for col in df.columns:
        if isinstance(col, str) and col.strip().lower() == 'audit leader':
            return col
    for col in df.columns:
        if isinstance(col, str) and 'audit leader' in col.lower():
            return col
    return None


def get_all_audit_leaders(workbook_path):
    print(f"Opening workbook to get audit leaders: {workbook_path}")
    all_leaders = set()
    excel_file = pd.ExcelFile(workbook_path)
    qa_sheets = [sheet for sheet in excel_file.sheet_names if is_qa_id_sheet(sheet)]
    print(f"Found {len(qa_sheets)} QA-ID sheets to process: {qa_sheets}")

    for sheet_name in qa_sheets:
        print(f"  Scanning sheet: {sheet_name} for audit leaders")
        try:
            header_row_idx = find_data_table_header_row(workbook_path, sheet_name)
            if header_row_idx is not None:
                df = pd.read_excel(excel_file, sheet_name=sheet_name,
                                   skiprows=header_row_idx, header=0)
                leader_col = find_audit_leader_column(df)
                if leader_col:
                    print(f"  Found leader column: {leader_col}")
                    leaders = df[leader_col].dropna().astype(str).unique()
                    leaders = [leader for leader in leaders if
                               leader.strip() and leader.strip().lower() != 'audit leader']
                    all_leaders.update(leaders)
                    print(f"  Added {len(leaders)} leaders: {', '.join(leaders)}")
                else:
                    print(f"  Could not find audit leader column in sheet {sheet_name}")
            else:
                print(f"  Could not find data table in sheet {sheet_name}")
        except Exception as e:
            print(f"  Error processing sheet {sheet_name}: {e}")
            traceback.print_exc()

    leader_list = sorted(list(all_leaders))
    print(f"Total unique audit leaders found: {len(leader_list)}")
    print(f"Leaders: {', '.join(leader_list)}")
    return leader_list


def normalize_leader_name(name):
    if not isinstance(name, str):
        return str(name).lower().strip() if name is not None else ""
    return name.lower().strip()


def filter_data_for_leader(df, leader_col, leader_name):
    if leader_col not in df.columns:
        return pd.DataFrame()
    normalized_leader = normalize_leader_name(leader_name)
    normalized_col = df[leader_col].fillna("").astype(str).apply(normalize_leader_name)
    filtered = df[normalized_col == normalized_leader].copy()
    if filtered.empty:
        filtered = df[normalized_col.str.contains(normalized_leader, na=False)].copy()
    return filtered


def copy_entire_sheet(source_sheet, target_sheet):
    try:
        print(f"    Copying entire sheet: {source_sheet.title}")
        for row in source_sheet.iter_rows():
            row_data = [cell.value for cell in row]
            target_sheet.append(row_data)
        for column in source_sheet.column_dimensions:
            target_sheet.column_dimensions[column].width = source_sheet.column_dimensions[column].width
        print(f"    Sheet copied successfully")
    except Exception as e:
        print(f"    Error copying sheet: {e}")
        traceback.print_exc()


def create_leader_workbook(input_path, output_path, leader_name):
    print(f"\nProcessing workbook for: {leader_name}")
    try:
        print(f"  Loading original workbook: {input_path}")
        original_wb = load_workbook(input_path)
        new_wb = Workbook()

        if "Sheet" in new_wb.sheetnames:
            del new_wb["Sheet"]

        found_data = False
        sheets_with_dnc = set()

        for sheet_name in original_wb.sheetnames:
            print(f"  Processing sheet: {sheet_name}")
            new_sheet = new_wb.create_sheet(title=sheet_name)

            if is_qa_id_sheet(sheet_name):
                print(f"    This is a QA-ID sheet - filtering for {leader_name}")
                has_dnc = False
                try:
                    header_row_idx = find_data_table_header_row(input_path, sheet_name)
                    if header_row_idx is not None:
                        print(f"    Found data table header at row {header_row_idx + 1}")
                        df = pd.read_excel(input_path, sheet_name=sheet_name,
                                           skiprows=header_row_idx, header=0)
                        leader_col = find_audit_leader_column(df)
                        if leader_col:
                            print(f"    Using column: {leader_col}")
                            filtered_df = filter_data_for_leader(df, leader_col, leader_name)
                            if not filtered_df.empty:
                                print(f"    Found {len(filtered_df)} rows for {leader_name}")
                                found_data = True
                                qa_col = next((col for col in filtered_df.columns
                                               if isinstance(col, str) and 'qa' in col.lower() and 'result' in col.lower()), None)
                                if qa_col:
                                    print(f"    Found QA Results column: {qa_col}")
                                    dnc_rows = filtered_df[filtered_df[qa_col].astype(str).str.lower() == 'dnc']
                                    if not dnc_rows.empty:
                                        print(f"    ⚠️ Found {len(dnc_rows)} DNC results in this sheet!")
                                        has_dnc = True
                                        sheets_with_dnc.add(sheet_name)
                                original_sheet = original_wb[sheet_name]
                                for row_idx in range(1, header_row_idx + 2):
                                    new_sheet.append([cell.value for cell in original_sheet[row_idx]])
                                for _, row in filtered_df.iterrows():
                                    new_sheet.append(row.tolist())
                                for cell in new_sheet[header_row_idx + 1]:
                                    cell.font = Font(bold=True)
                                for i in range(1, len(row) + 1):
                                    new_sheet.column_dimensions[get_column_letter(i)].width = 15
                            else:
                                print(f"    No data found for {leader_name} in sheet {sheet_name}")
                                original_sheet = original_wb[sheet_name]
                                for row_idx in range(1, min(header_row_idx + 10, len(list(original_sheet.rows))) + 1):
                                    new_sheet.append([cell.value for cell in original_sheet[row_idx]])
                        else:
                            print(f"    Could not find audit leader column")
                            original_sheet = original_wb[sheet_name]
                            for row_idx in range(1, min(20, len(list(original_sheet.rows))) + 1):
                                new_sheet.append([cell.value for cell in original_sheet[row_idx]])
                    else:
                        print(f"    Could not find data table header")
                        original_sheet = original_wb[sheet_name]
                        for row_idx in range(1, min(20, len(list(original_sheet.rows))) + 1):
                            new_sheet.append([cell.value for cell in original_sheet[row_idx]])
                except Exception as e:
                    print(f"    Error processing QA-ID sheet {sheet_name}: {e}")
                    traceback.print_exc()
                    original_sheet = original_wb[sheet_name]
                    for row_idx in range(1, min(15, len(list(original_sheet.rows))) + 1):
                        try:
                            new_sheet.append([cell.value for cell in original_sheet[row_idx]])
                        except:
                            break
            else:
                print(f"    This is not a QA-ID sheet - copying entirely")
                try:
                    original_sheet = original_wb[sheet_name]
                    copy_entire_sheet(original_sheet, new_sheet)
                except Exception as e:
                    print(f"    Error copying non-QA-ID sheet {sheet_name}: {e}")
                    traceback.print_exc()

        for sheet_name in new_wb.sheetnames:
            sheet = new_wb[sheet_name]
            if is_qa_id_sheet(sheet_name):
                sheet.sheet_properties.tabColor = "FF0000" if sheet_name in sheets_with_dnc else "00FF00"
                print(f"  {'🔴' if sheet_name in sheets_with_dnc else '🟢'} Tab for QA-ID sheet {sheet_name} colored")

        print(f"  Saving workbook to: {output_path}")
        new_wb.save(output_path)
        return found_data

    except Exception as e:
        print(f"Error creating workbook for {leader_name}: {e}")
        traceback.print_exc()
        return False


def split_audit_data(input_file, output_dir='audit_leaders'):
    try:
        if not os.path.exists(output_dir):
            print(f"Creating output directory: {output_dir}")
            os.makedirs(output_dir)
        else:
            print(f"Output directory already exists: {output_dir}")

        audit_leaders = get_all_audit_leaders(input_file)
        print(f"Found {len(audit_leaders)} audit leaders across QA-ID sheets")

        success_count = 0
        for leader in audit_leaders:
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
        import argparse

        parser = argparse.ArgumentParser(description="Split audit data Excel file by audit leaders.")
        parser.add_argument("-i", "--input", help="Path to input Excel file")
        parser.add_argument("-o", "--output", help="Path to output directory")
        args = parser.parse_args()

        input_file = args.input or input("Enter path to Excel file: ").strip().strip('"\'')
        if not os.path.exists(input_file):
            print(f"ERROR: Input file not found at {input_file}")
            sys.exit(1)

        output_dir = args.output
        if not output_dir:
            output_dir = input("Enter output directory path (press Enter for default 'audit_leaders'): ").strip().strip('"\'')
            if not output_dir:
                output_dir = os.path.join(os.path.dirname(input_file), "audit_leaders")
                print(f"Using default output directory: {output_dir}")

        if split_audit_data(input_file, output_dir):
            print("Script completed successfully!")
        else:
            print("Script completed with errors!")
    except Exception as e:
        print(f"ERROR: {str(e)}")
        traceback.print_exc()
