"""
Compare two Excel files cell-by-cell to see differences
"""

from pathlib import Path
from loguru import logger
import openpyxl
import sys

def compare_excel_files(file1_path: str, file2_path: str):
    """Compare two Excel files cell by cell"""
    
    path1 = Path(file1_path).resolve()
    path2 = Path(file2_path).resolve()
    
    if not path1.exists():
        logger.error(f"File 1 does not exist: {path1}")
        return
    
    if not path2.exists():
        logger.error(f"File 2 does not exist: {path2}")
        return
    
    logger.info("="*70)
    logger.info("EXCEL FILE COMPARISON")
    logger.info("="*70)
    logger.info(f"File 1: {path1.name}")
    logger.info(f"File 2: {path2.name}")
    logger.info("="*70)
    
    try:
        wb1 = openpyxl.load_workbook(path1, read_only=True, data_only=True)
        wb2 = openpyxl.load_workbook(path2, read_only=True, data_only=True)
        
        # Compare sheet names
        sheets1 = set(wb1.sheetnames)
        sheets2 = set(wb2.sheetnames)
        
        if sheets1 != sheets2:
            logger.warning("Sheet names differ!")
            logger.info(f"  File 1 sheets: {sheets1}")
            logger.info(f"  File 2 sheets: {sheets2}")
        else:
            logger.success(f"✓ Both files have same sheets: {sheets1}")
        
        # Compare common sheets
        common_sheets = sheets1 & sheets2
        
        total_differences = 0
        
        for sheet_name in common_sheets:
            logger.info(f"\n{'='*70}")
            logger.info(f"Comparing sheet: '{sheet_name}'")
            logger.info(f"{'='*70}")
            
            sheet1 = wb1[sheet_name]
            sheet2 = wb2[sheet_name]
            
            max_row = max(sheet1.max_row, sheet2.max_row)
            max_col = max(sheet1.max_column, sheet2.max_column)
            
            differences_in_sheet = 0
            
            for row in range(1, min(max_row + 1, 100)):  # Check first 100 rows
                for col in range(1, min(max_col + 1, 26)):  # Check first 26 columns (A-Z)
                    cell1 = sheet1.cell(row, col)
                    cell2 = sheet2.cell(row, col)
                    
                    val1 = cell1.value
                    val2 = cell2.value
                    
                    if val1 != val2:
                        differences_in_sheet += 1
                        total_differences += 1
                        
                        if differences_in_sheet <= 20:  # Show first 20 differences
                            col_letter = openpyxl.utils.get_column_letter(col)
                            logger.warning(f"  Cell {col_letter}{row}:")
                            logger.warning(f"    File 1: {val1}")
                            logger.warning(f"    File 2: {val2}")
            
            if differences_in_sheet == 0:
                logger.success(f"✓ No differences in sheet '{sheet_name}'")
            else:
                logger.error(f"❌ Found {differences_in_sheet} differences in sheet '{sheet_name}'")
        
        wb1.close()
        wb2.close()
        
        logger.info(f"\n{'='*70}")
        if total_differences == 0:
            logger.success("✓✓✓ FILES ARE IDENTICAL!")
        else:
            logger.error(f"❌❌❌ Found {total_differences} total differences")
        logger.info(f"{'='*70}")
        
    except Exception as e:
        logger.error(f"Error comparing files: {e}")


if __name__ == "__main__":
    if len(sys.argv) < 3:
        print("Usage: python compare_excel_cells.py <file1> <file2>")
        print("\nExample:")
        print('python compare_excel_cells.py "data/debug_uploads/file1.xlsx" "data/excel_files/file2.xlsx"')
        sys.exit(1)
    
    file1 = sys.argv[1]
    file2 = sys.argv[2]
    compare_excel_files(file1, file2)
