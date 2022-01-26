"""
Copyright 2021 Rob Randall (rob.randall1@gmail.com)

Permission is hereby granted, free of charge, to any person obtaining a copy
of this software and associated documentation files (the "Software"), to
deal in the Software without restriction, including without limitation the
rights to use, copy, modify, merge, publish, distribute, sublicense, and/or
sell copies of the Software, and to permit persons to whom the Software is
furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in
all copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS
OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING
FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS
IN THE SOFTWARE.
"""
import argparse
import sys
from typing import Sequence

import numpy as np
import pandas as pd
from openpyxl.utils.cell import get_column_letter


class colors:
    HEADER = '\033[95m'
    BLUE = '\033[34m'
    CYAN = '\033[36m'
    YELLOW = '\033[33m'
    GREEN = '\033[32m'
    RED = '\033[31m'
    FAIL = '\033[91m'
    ENDC = '\033[0m'
    BOLD = '\033[1m'
    UNDERLINE = '\033[4m'


def blue(str):
    return f'{colors.BLUE}{str}{colors.ENDC}'


def cyan(str):
    return f'{colors.CYAN}{str}{colors.ENDC}'


def green(str):
    return f'{colors.GREEN}{str}{colors.ENDC}'


def red(str):
    return f'{colors.RED}{str}{colors.ENDC}'


def bold(str):
    return f'{colors.BOLD}{str}{colors.ENDC}'


def any(str):
    return f'\033[1m{str}{colors.ENDC}'


def excel_diff(file1: str, file2: str) -> int:
    """Diff the sheets in 2 excel files.
    file1 is considered the 'old' version, and file2 the 'new' version."""

    diff = 0

    file1_sheets = pd.read_excel(file1, sheet_name=None, header=None, index_col=None)
    file1_sheet_names = set(file1_sheets.keys())
    file2_sheets = pd.read_excel(file2, sheet_name=None, header=None, index_col=None)
    file2_sheet_names = set(file2_sheets.keys())

    for sheet_name in file1_sheet_names.difference(file2_sheet_names):
        diff = 1
        print(f'Sheet {cyan(sheet_name)} is only in {red(file1)}')

    for sheet_name in file2_sheet_names.difference(file1_sheet_names):
        diff = 1
        print(f'Sheet {cyan(sheet_name)} is only in {green(file2)}')

    for sheet_name in file1_sheet_names.intersection(file2_sheet_names):
        df1 = file1_sheets[sheet_name]
        df2 = file2_sheets[sheet_name]

        if df1.shape != df2.shape:
            diff = 1
            print(f'Sheet {cyan(sheet_name)} has a different shape: {red(df1.shape)} --> {green(df2.shape)}')

        else:

            comparison_values = df1.values == df2.values

            df_equals = df1.equals(df2)

            if not df_equals:
                diff = 1
                df_compare = df1.compare(df2, keep_shape=True)

                rows, cols = np.where(comparison_values == False)
                for r, c in zip(rows, cols):
                    diff_cell = df_compare.iloc[r][c]
                    if not diff_cell.isnull().all():
                        print(
                            f'Sheet {cyan(sheet_name)} cell {blue(get_column_letter(c+1))}{blue(r+1)} has changed: {red(diff_cell["self"])} --> {green(diff_cell["other"])}'
                        )
    return diff


def git_diff(argv: Sequence[str] = None) -> int:
    """Support arguments used by git diff to specify the excel files to diff"""
    parser = argparse.ArgumentParser(description='Excel Git diff')
    parser.add_argument('path')
    parser.add_argument('old_file', metavar='old-file')
    parser.add_argument('old-hex')
    parser.add_argument('old-mode')
    parser.add_argument('new_file', metavar='old-file')
    parser.add_argument('new-hex')
    parser.add_argument('new-mode')
    args = parser.parse_args(argv)
    excel_diff(args.old_file, args.new_file)
    return 0  # git does not like an error return


def diff(argv: Sequence[str]) -> int:
    """Diff to excel files specified on the command line. Decide between git arg set and simple
    2 file name args"""
    # check if this has 8 args, and if so treat as being called by git diff
    if len(argv) == 8:
        return git_diff(argv[1:])
    else:
        # Not git diff, so 2 file args expected
        parser = argparse.ArgumentParser(description='Excel Git diff')
        parser.add_argument('file1')
        parser.add_argument('file2')
        args = parser.parse_args(argv[1:])
        return excel_diff(args.file1, args.file2)


def main(argv: Sequence[str]) -> int:
    return diff(argv)


if __name__ == '__main__':
    exit(main(sys.argv))
