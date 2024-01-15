from openpyxl import Workbook
from openpyxl import load_workbook
from argparse import ArgumentParser
from argparse import RawTextHelpFormatter
import os
import sys
import platform
from pathlib import Path
import pprint
import numpy as np
from openpyxl.styles import NamedStyle, Font, Border, Side, Color, PatternFill, Alignment
from openpyxl.utils import get_column_letter


def parse_arg():
    parser = ArgumentParser(description="Generate peer review files.",formatter_class=RawTextHelpFormatter,
                            epilog=     "Example：\n"
                                        "Excel：generate.py src.xlsx\n"
                                        )
    parser.add_argument("src",
                        help="Read information from this excel file", metavar="Input-File", type=str)
    args = parser.parse_args()
    return vars(args)


if __name__ == '__main__':
    args = parse_arg()
    src = args["src"]
    print(f"Start generating peer review files from {src}...")
    #load src.xlsx
    wb = load_workbook(src)
    ws = wb['columns']
    cols = []
    for row in ws.iter_rows(min_row=2, max_col=3, max_row=ws.max_row):
        col = []
        cols.append(col)
        for cell in row:
            col.append(cell.value)
    cols = np.array(cols)
    print("Columns:")
    pprint.pprint(cols)

    for ws in wb.worksheets:
        if ws.title != "columns":
            students = []
            for row in ws.iter_rows(min_row=2, max_col=3, max_row=ws.max_row):
                student = []
                students.append(student)
                for cell in row:
                    student.append(cell.value)
            print(f"Students in {ws.title}:")
            pprint.pprint(students)
            newwb= Workbook()
            newws = newwb.active
            newws.title = f"Peer Review for Team {ws.title}"
            row = list(map(lambda x: f"{x[0]}({x[1]} to {x[2]})", cols.tolist()))
            row.insert(0,"Student Name")
            newws.append(row)
            for student in students:
                row = []
                row.append(f"{student[0]}, {student[1]}")
                for i in range(0, len(cols)):
                    row.append(0)
                newws.append(row)



            for row in newws.iter_rows():
                for cell in row:
                    if cell.col_idx == 1 or cell.row==1:
                        cell.fill = PatternFill(start_color='CCCCCC', end_color='CCCCCC', fill_type="solid")
                    else:
                        cell.fill = PatternFill(start_color='00CC00', end_color='00CC00', fill_type="solid")
                        cell.number_format = '0'

            newws.column_dimensions['A'].width = 30
            for col in newws.iter_cols(min_col=2, max_col=newws.max_column):
                col[0].alignment = Alignment(horizontal='center', vertical='center')
                newws.column_dimensions[get_column_letter(col[0].column)].width = 20


            newwb.save(f"Peer Review for Team {ws.title}.xlsx")
            newwb.close()

    #read col tab

    #for each rest of tabs(teams), generate excel files

    print("Complete generating peer review files...")