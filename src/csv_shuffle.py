import copy
import csv
import os
import string
import sys

from typing import List, Tuple, Union

DATA_PATH = '/home/adam/tmp/csv-data'
# IN_FILE = 'sample01.csv'
IN_FILE = 'InventorySearchResults.csv'
# OUT_FILE = 'sample01-shuffle.csv'
OUT_FILE = 'InventorySearchResults-shuffle.csv'
ENCODING = 'utf-8'
ENCODING_ERRORS = 'backslashreplace'
HEADERS_DEFINED = False
OUTPUT_COLUMNS = ['G',
                  'A',
                  'B',
                  'H',
                  'E']
# OUTPUT_HEADERS = ['Column 0 (A)',
#                   'Column 7 (H)',
#                   'Column 4 (E)',
#                   'Column 5 (F)',
#                   'Column 2 (C)']
OUTPUT_HEADERS = ['asset_id',
                  'tsnumber',
                  'serialnumber',
                  'Model',
                  'datecreated',
                  'assetGUID']
OUTPUT_INDEXES = [5,
                  3,
                  1,
                  2]

IN_PATH = os.path.join(DATA_PATH, IN_FILE)
OUT_PATH = os.path.join(DATA_PATH, OUT_FILE)


def _column_to_index(column: str) -> int:
    """
    Convert a spreadsheet column letter to a numeric index.

    Notes:
        Pulled from a stack overflow answer.
        Made a few minor changes.  First I change the name to be more in
        line with the naming preference I prefer.  Second I change it so the
        index number returned is 0 based because I want to use it to pull
        from a 0 based list.

    :param column:
        The string representation of the columns like
        'A', 'B', ... 'Z', 'AA', 'AB', ... 'AZ', 'BA'...
    :return:
        A number that will correspond to the array index of a column letter
        where 'A' -> 0, 'B' -> 1 and so forth.
    """
    # TODO: What if someone sends unicode???
    num = 0
    for c in column:
        if c in string.ascii_letters:
            num = (num * 26) + (ord(c.upper()) - ord('A')) + 1
    return num - 1


def _calculate_output_indexes(
        indexes: List[int] = None,
        columns: List[str] = None,
        in_headers: List[str] = None,
        out_headers: List[str] = None) -> Tuple[bool,
                                                Union[str, None],
                                                List[int]]:
    """
    Translate the provided list of columns to indexes in the rows of the
    data set.

    Notes:
        The expectation is that one of the three different lists will be
        provided.  If more than one is provided the one closest to the
        beginning of the parameter list will take precedence.  If none is
        provided, an empty list will be returned.

    :param indexes:
        This is the case where what is provided is already what we want so
        a copy of the list is returned.
    :param columns:
        A list of the column letters of the database that look like
        'A', 'B', ... 'Z', 'AA', 'AB', ... 'AZ', 'BA'... and so forth.
    :param in_headers:
        The list of column headers in the original data set in the order
        that they appear.  These need to be provided with the out_headers
        parameter.
    :param out_headers:
        The headers that will appear in the shuffled output in the order that
        they will appear.  These will be translated to index numbers by
        finding the elements in the in_headers list.  The comparisons will be
        for exact (case-sensitive) matches.  If the same header column appears
        in in_headers more than once, the first one identified will always be
        used to identify the translation index.  If an element of this list
        can't be found in in_headers it will be identified as a failure and
        skipped.
    :return:
        A tuple containing the following elements:
            0 - flag to indicate if the translation succeded
            1 - string describing cause of failure if translation failed
            2 - list of translated values.  This may not contain entries for
                all provided values since some may fail.  If any fail the
                whole translation will be considered a failure but what could
                be translated will be returned.
    """
    cof = None,
    out_indexes = []

    if (indexes is not None) and (len(indexes) > 0):
        success = True,
        out_indexes = copy.copy(indexes)
    elif (columns is not None) and (len(columns) > 0):
        for column in columns:
            out_indexes.append(_column_to_index(column))
        success = True,
    elif (out_headers is not None) and (len(out_headers) > 0):
        if in_headers is not None:
            unmatched_headers = []
            for out_header in out_headers:
                out_index = None
                for in_index, in_header in enumerate(in_headers):
                    if out_header == in_header:
                        out_index = in_index
                        break
                if out_index is not None:
                    out_indexes.append(out_index)
                else:
                    unmatched_headers.append(out_header)
            if len(unmatched_headers) == 0:
                success = True
            else:
                success = False
                cof = (f'some headers {unmatched_headers} matched nothing in '
                       f'original headers {in_headers}')
        else:
            success = False
            cof = 'in_headers must be provided with out_headers'
    else:
        success = False
        cof = 'no input indexes provided'

    return success, cof, out_indexes


def main() -> None:
    try:
        with open(IN_PATH,
                  'r',
                  encoding=ENCODING,
                  errors=ENCODING_ERRORS) as in_fh:
            csv_reader = csv.reader(in_fh,
                                    delimiter=',',
                                    quotechar='"')
            data_rows_in = []
            for row in csv_reader:
                data_rows_in.append(row)
    except IOError as ioe:
        raise RuntimeError(f'failed to read lines from IN_PATH ({IN_PATH}: '
                           f'{ioe}')

    output_indexes = _calculate_output_indexes(
        in_headers=copy.copy(data_rows_in[0]),
        out_headers=OUTPUT_HEADERS)[2]
    try:
        with open(OUT_PATH,
                  'w',
                  encoding=ENCODING,
                  errors=ENCODING_ERRORS) as out_fh:
            csv_writer = csv.writer(out_fh,
                                    delimiter=',',
                                    quotechar='"')
            for row in data_rows_in:
                data_row_out = []
                for index in output_indexes:
                    data_row_out.append(row[index])
                csv_writer.writerow(data_row_out)
    except IOError as ioe:
        raise RuntimeError(f'failed to write lines to OUT_PATH ({OUT_PATH}: '
                           f'{ioe}')


if __name__ == '__main__':
    try:
        main()
    except RuntimeError as rte:
        print(f'shuffle failed: {rte}')
        sys.exit(1)
