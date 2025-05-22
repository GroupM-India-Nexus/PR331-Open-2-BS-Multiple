from openpyxl.styles import Border, Side

def apply_borders(out_worksheet, start_row, start_col, end_row, end_col):
    # Define border styles : thin
    medium_border = Border(
        left=Side(style='medium'),
        right=Side(style='medium'),
        top=Side(style='medium'),
        bottom=Side(style='medium')
    )
    
    for row in range(start_row, end_row + 1):
            for col in range(start_col, end_col + 1):
                out_worksheet.cell(row=row, column=col).border = medium_border
                
    return out_worksheet        