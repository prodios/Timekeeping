from openpyxl.utils import get_column_letter


def create_cell(cell, **kwargs):
    for attr, value in kwargs.items():
        if value != None:
            setattr(cell, attr, value)


def fix_width(worksheet):
    for col in worksheet.columns:
        max_length = 2
        column = col[0].column
        for cell in col:
            try:
                if len(str(cell.value)) > max_length:
                    max_length = len(cell.value)
            except TypeError:
                pass
        adjusted_width = (max_length + 2) * 1.2
        worksheet.column_dimensions[get_column_letter(column)].width = adjusted_width
