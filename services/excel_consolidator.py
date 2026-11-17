import os
import copy

# NumPy 2.x & OpenPyXL compatibility polyfill
try:
    import numpy as _np
    for _alias, _target in [
        ('short', getattr(_np, 'int16', None)),
        ('ushort', getattr(_np, 'uint16', None)),
        ('int_', getattr(_np, 'int64', getattr(_np, 'int32', None))),
        ('uint', getattr(_np, 'uint64', getattr(_np, 'uint32', None))),
        ('longlong', getattr(_np, 'int64', None)),
        ('ulonglong', getattr(_np, 'uint64', None)),
        ('half', getattr(_np, 'float16', None)),
        ('single', getattr(_np, 'float32', None)),
        ('double', getattr(_np, 'float64', None)),
        ('longdouble', getattr(_np, 'float64', None)),
        ('intp', getattr(_np, 'int64', getattr(_np, 'int32', None))),
        ('uintp', getattr(_np, 'uint64', getattr(_np, 'uint32', None))),
    ]:
        if _target is not None and not hasattr(_np, _alias):
            setattr(_np, _alias, _target)
except Exception:
    pass

from openpyxl import load_workbook, Workbook
from openpyxl.utils import get_column_letter, column_index_from_string
from utils import apply_workbook_watermark

def get_unique_sheets(selected_files):
    """
    Reads the unique sheet names across all selected files.
    """
    unique_sheets = set()
    for file in selected_files:
        wb = load_workbook(file, read_only=True)
        unique_sheets.update(wb.sheetnames)
        wb.close()
    return sorted(unique_sheets)

def consolidate(selected_files, unique_sheets, save_path, progress_callback=None):
    """
    Consolidates the selected Excel files into a single file at save_path,
    retaining formatting, sheet names, row dimensions, and column widths.
    """
    target_wb = Workbook()
    target_wb.remove(target_wb.active)  # Remove default active sheet
    
    for sheet_name in unique_sheets:
        target_wb.create_sheet(sheet_name)
    
    for i, file_path in enumerate(selected_files):
        file_name = os.path.basename(file_path)
        source_wb = load_workbook(file_path)
        
        for sheet_name in source_wb.sheetnames:
            source_sheet = source_wb[sheet_name]
            
            if sheet_name not in target_wb.sheetnames:
                continue
            target_sheet = target_wb[sheet_name]
            
            # Determine starting row for appending data
            if target_sheet.max_row == 1 and target_sheet.cell(1, 1).value is None:
                last_row = 5  # Leave some blank rows at the start if new
            else:
                last_row = target_sheet.max_row
            
            if last_row > 5:
                last_row += 1  # Add a blank separating row
            
            data_start_row = last_row + 1
            
            # Copy cells
            for r_idx, row in enumerate(source_sheet.iter_rows(values_only=False), 1):
                target_row = data_start_row + r_idx - 1

                # Column A is reserved for the source file name
                target_sheet.cell(row=target_row, column=1, value=file_name)
                
                for c_idx, cell in enumerate(row, 1):
                    target_cell = target_sheet.cell(
                        row=target_row,
                        column=c_idx + 1,  # Shift data columns to the right by 1
                        value=cell.value
                    )
                    if cell.has_style:
                        target_cell.font = copy.copy(cell.font)
                        target_cell.border = copy.copy(cell.border)
                        target_cell.fill = copy.copy(cell.fill)
                        target_cell.number_format = cell.number_format
                        target_cell.alignment = copy.copy(cell.alignment)
            
            # Recreate merged cells
            for merged_range in source_sheet.merged_cells.ranges:
                new_min_row = merged_range.min_row + data_start_row - 1
                new_max_row = merged_range.max_row + data_start_row - 1
                new_min_col = merged_range.min_col + 1  # Shifted column index
                new_max_col = merged_range.max_col + 1  # Shifted column index
                
                target_sheet.merge_cells(
                    start_row=new_min_row,
                    start_column=new_min_col,
                    end_row=new_max_row,
                    end_column=new_max_col
                )
            
            # Copy row heights
            for row_idx, dim in source_sheet.row_dimensions.items():
                if dim.height is not None:
                    target_sheet.row_dimensions[row_idx + data_start_row - 1].height = dim.height
            
            # Copy column widths
            for col_letter, dim in source_sheet.column_dimensions.items():
                if dim.width is not None:
                    col_idx = column_index_from_string(col_letter)
                    new_col_idx = col_idx + 1
                    new_col_letter = get_column_letter(new_col_idx)
                    target_sheet.column_dimensions[new_col_letter].width = dim.width
        
        source_wb.close()
        
        if progress_callback:
            progress_callback((i + 1) / len(selected_files))
            
    # Auto-size the filename column A
    for sheet in target_wb.worksheets:
        sheet.column_dimensions['A'].width = 30

    # Apply official author metadata
    apply_workbook_watermark(target_wb)

    target_wb.save(save_path)
    target_wb.close()
