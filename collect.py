from openpyxl import Workbook
from openpyxl import load_workbook
from argparse import ArgumentParser
from argparse import RawTextHelpFormatter
import os
import sys
import platform
from pathlib import Path
import pprint
def parse_arg():
    parser = ArgumentParser(description="Collect peer review files.",formatter_class=RawTextHelpFormatter,
                            epilog=     "Example：\n"
                                        "Excel：collect.py folder src.xlsx dest.xlsx\n"
                                        )
    parser.add_argument("folder",
                        help="Read peer review excel files from this folder", metavar="Input-File", type=str)
    parser.add_argument("src",
                        help="Read information from this excel file", metavar="Input-File", type=str)
    parser.add_argument("dst",
                        help="Write results to this excel file", metavar="Output-File", type=str)

    args = parser.parse_args()
    return vars(args)

if __name__ == '__main__':
    print("Start collecting peer review files...")


    print("Complete collecting peer review files...")