import configparser
import copy
import csv
import os
import string
import sys
from typing import Dict, List, Tuple, Union

# Values for data read and write parameters that will be used if not
# defined in run time parameters.
import openpyxl

DEFAULT_CHARACTER_ENCODING = 'utf-8'
DEFAULT_ENCODING_ERRORS = 'backslashreplace'


def _column_to_index(column: str) -> int:
    """
    Convert a spreadsheet column letter to a numeric index.

    Notes:
        Pulled from a stack overflow answer.
        Made a few minor changes.  First I change the name to be more in
        line with the naming preference I prefer.  Second I change it so the
        index number returned is 0 based because I want to use it to pull
        from a 0 based list.

        It is assumed that the column str is composed of ASCII letters.
        If that is not the case, the behavior is not defined.
        So don't send something bizarre like a smiley face.

    :param column:
        The string representation of the columns like
        'A', 'B', ... 'Z', 'AA', 'AB', ... 'AZ', 'BA'...
    :return:
        A number that will correspond to the array index of a column letter
        where 'A' -> 0, 'B' -> 1 and so forth.
    """
    num = 0
    for col_letter in column:
        if col_letter in string.ascii_letters:
            num = (num * 26) + (ord(col_letter.upper()) - ord('A')) + 1
    return num - 1


def _calculate_output_indexes(in_headers: List[str],
                              params: dict,) -> Tuple[bool,
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

    :param in_headers:
        The list of column headers in the original data set in the order
        that they appear.  These need to be provided with the out_headers
        parameter.
    :param params:
        Collection of runtime parameters for the script that will be searched
        for valid column definitions.
    :return:
        A tuple containing the following elements:
            0 - flag to indicate if the translation succeded
            1 - string describing cause of failure if translation failed
            2 - list of translated values.  This may not contain entries for
                all provided values since some may fail.  If any fail the
                whole translation will be considered a failure but what could
                be translated will be returned.
    """
    cof = None
    out_indexes = []

    if ('column_indexes' in params) and (len(params['column_indexes']) > 0):
        success = True
        out_indexes = copy.copy(params['column_indexes'])
    elif ('column_letters' in params) and (len(params['column_letters']) > 0):
        for column in params['column_letters']:
            out_indexes.append(_column_to_index(column))
        success = True
    elif ('column_headers' in params) and (len(params['column_headers']) > 0):
        if in_headers is not None:
            unmatched_headers = []
            for out_header in params['column_headers']:
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


def _read_config(config_path: str) -> Tuple[bool,
                                            Union[str, None],
                                            Union[dict, None]]:
    """
    Verify that config_path is a valid file and read its contents into a
    ConfigParser object.

    :param config_path:
        Path to file containing the configuration parameters for this script.
    :return:
        A tuple containing the following elements:
            0 - flag to indicate if the configuration parameters were read
                and validated
            1 - string describing reason configuration parameter read
                failed
            2 - Dictionary containing the configuration parameters needed
                to run this script.  This may not be a complete set of
                parameters if the validation is not found to be valid.
    """
    cof = None

    config_path_good = True
    if config_path is None:
        config_path_good = False
        cof = 'configuration file must be provided'
    elif ((not os.path.isfile(config_path)) or
          (not os.access(config_path, os.R_OK))):
        config_path_good = False
        cof = f'configuration file ({config_path}) must be a readable file'

    if not config_path_good:
        return config_path_good, cof, None

    config_parser = configparser.ConfigParser()
    try:
        with open(config_path,
                  'r',
                  encoding=DEFAULT_CHARACTER_ENCODING,
                  errors=DEFAULT_ENCODING_ERRORS) as config_fh:
            config_parser.read_file(config_fh)
    except IOError as exc:
        return False, \
               f'failed to read configuration file ({config_path}): {exc}', \
               None

    return _validate_config(config_parser)


def _read_csv_data(params: dict) -> List[List[str]]:
    """
    Read the rows of data in the CSV data file defined in params.

    :param params:
        Collection of runtime parameters provided by the caller that have
        been verified and translated to a dictionary.
    :return:
        A list of the data rows in the input data file in the order that
        they appear in the original data file.
    :raise RuntimeError:
        If there is a problem opening or reading the input file specified
        in params as a CSV data file this exception will be raised.
    """
    data_rows = []
    try:
        with open(params['input_file_path'],
                  'r',
                  encoding=params['character_encoding'],
                  errors=params['character_encoding_errors']) as in_fh:
            csv_reader = csv.reader(in_fh,
                                    delimiter=',',
                                    quotechar='"')
            for row in csv_reader:
                data_rows.append(row)
    except IOError as ioe:
        raise RuntimeError(f'failed to read lines from input data file '
                           f'({params["input_file_path"]}): {ioe}')

    return data_rows


def _read_xlsx_data(params: dict) -> List[List[str]]:
    """
    Read the rows of data in the Excel spreadsheet defined in params.

    :param params:
        Collection of runtime parameters provided by the caller that have
        been verified and translated to a dictionary.
    :return:
        A list of the data rows in the input data file in the order that
        they appear in the original data file.
    :raise RuntimeError:
        If there is a problem opening or reading the input file specified
        in params as an Excel spreadsheet this exception will be raised.
    """
    data_rows = []
    try:
        wb = openpyxl.load_workbook(params['input_file_path'],
                                    read_only=True,
                                    data_only=True)
        if params['input_sheet_name'] not in wb.sheetnames:
            raise RuntimeError(f'no sheet named '
                               f'({params["input_sheet_name"]}) is contained '
                               f'in input data file '
                               f'({params["input_file_path"]})')
        ws = wb[params['input_sheet_name']]

        for row in ws.values:
            data_rows.append(list(row))
    except IOError as ioe:
        raise RuntimeError(f'failed to read lines from input data file '
                           f'({params["input_file_path"]}): {ioe}')

    return data_rows


def _validate_config(config_raw: configparser.ConfigParser) -> \
        Tuple[bool, Union[str, None], Dict]:
    """
    Check config_parser for the parameters needed to operate this script.

    :param config_raw:
        The raw configuration data provided to this run of the script.
    :return:
        A tuple containing the following elements:
            0 - flag to indicate if config_parser contains the expected
                parameters and that they are valid
            1 - string describing issues that indicate that config_parser
                is not valid for running this script
            2 - Dictionary containing the parameters needed to run this
                script.  Some parameters may be defined even if the overall
                validation is considered a failure.
    """
    invalid = []
    config_map = {}

    if not config_raw.has_section('data_files'):
        invalid.append('data_files section not defined')
    else:
        if config_raw.has_option('data_files', 'input_path'):
            input_path = config_raw.get('data_files', 'input_path')
            if not (os.path.isdir(input_path)
                    and os.access(input_path, os.R_OK)):
                invalid.append(
                    'data_files.input_path is not a readable directory')
        else:
            input_path = None
            invalid.append('data_files.input_path not defined')

        if config_raw.has_option('data_files', 'input_file_name'):
            input_file_name = config_raw.get('data_files', 'input_file_name')
        else:
            input_file_name = None
            invalid.append('data_files.input_file_name not defined')

        if config_raw.has_option('data_files', 'input_file_extension'):
            input_file_extension = config_raw.get('data_files',
                                                  'input_file_extension')
            config_map['input_data_type'] = input_file_extension
        else:
            input_file_extension = None
            invalid.append('data_files.input_file_extension not defined')

        if ((input_path is not None) and
                (input_file_name is not None) and
                (input_file_extension is not None)):
            input_file_path = os.path.join(os.path.abspath(input_path),
                                           f'{input_file_name}.'
                                           f'{input_file_extension}')
            if not (os.path.isfile(input_file_path)
                    and os.access(input_file_path, os.R_OK)):
                invalid.append(f'input_file_path ({input_file_path} is not '
                               f'a readable file')
            else:
                config_map['input_file_path'] = input_file_path

        if config_raw.has_option('data_files', 'input_sheet_name'):
            config_map['input_sheet_name'] = config_raw.get('data_files',
                                                            'input_sheet_name')
        elif input_file_extension == 'xlsx':
            invalid.append('data_files.input_sheet_name must be defined '
                           'when data_files.input_file_extension is xlsx')

        if config_raw.has_option('data_files', 'output_path'):
            output_path = config_raw.get('data_files', 'output_path')
            if not (os.path.isdir(output_path)
                    and os.access(output_path, os.W_OK)):
                invalid.append(
                    'data_files.output_path is not a writeable directory')
        else:
            output_path = None
            invalid.append('data_files.output_path not defined')

        if config_raw.has_option('data_files', 'output_file_name'):
            output_file_name = config_raw.get('data_files', 'output_file_name')
        else:
            output_file_name = None
            invalid.append('data_files.output_file_name not defined')

        if config_raw.has_option('data_files', 'output_file_extension'):
            output_file_extension = config_raw.get('data_files',
                                                   'output_file_extension')
        else:
            output_file_extension = None
            invalid.append('data_files.output_file_extension not defined')

        if ((output_path is not None) and
                (output_file_name is not None) and
                (output_file_extension is not None)):
            config_map['output_file_path'] = os.path.join(
                os.path.abspath(output_path),
                f'{output_file_name}.{output_file_extension}')

        if config_raw.has_option('data_files', 'character_encoding'):
            config_map['character_encoding'] = config_raw.get(
                'data_files',
                'character_encoding')
        else:
            config_map['character_encoding'] = DEFAULT_CHARACTER_ENCODING

        if config_raw.has_option('data_files', 'character_encoding_errors'):
            config_map['character_encoding_errors'] = config_raw.get(
                'data_files',
                'character_encoding_errors')
        else:
            config_map['character_encoding_errors'] = DEFAULT_ENCODING_ERRORS

    if not config_raw.has_section('data_columns'):
        invalid.append('data_columns section not defined')
    else:
        columns_defined = False

        if config_raw.has_option('data_columns', 'column_headers'):
            headers = config_raw.get('data_columns', 'column_headers')
            config_map['column_headers'] = [col.strip()
                                            for col in headers.splitlines()]
            columns_defined = True

        if config_raw.has_option('data_columns', 'column_letters'):
            letters = config_raw.get('data_columns', 'column_letters')
            config_map['column_letters'] = [col.strip()
                                            for col in letters.splitlines()]
            columns_defined = True

        if config_raw.has_option('data_columns', 'column_indexes'):
            indexes = config_raw.get('data_columns', 'column_indexes')
            config_map['column_indexes'] = [int(col.strip())
                                            for col in indexes.splitlines()]
            columns_defined = True

        if not columns_defined:
            invalid.append('one of column headers, letters, or indexes must'
                           'be defined')

    if len(invalid) > 0:
        valid = False
        invalid_str = (f'one or more invalid configuration parameters '
                       f'identified: {", ".join(invalid)}')
    else:
        valid = True
        invalid_str = None

    return valid, invalid_str, config_map


def main(params: dict) -> None:
    """
    Read the input data file and write the data columns specified in a new
    file.

    :param params:
        Collection of runtime parameters provided by the caller that have
        been verified and translated to a dictionary.
    :raise RuntimeError:
        Any expected exceptions will be redefined and raised in this form.
    """
    if params['input_data_type'].lower() == 'csv':
        data_rows_in = _read_csv_data(params)
    elif params['input_data_type'].lower() == 'xlsx':
        data_rows_in = _read_xlsx_data(params)
    else:
        raise RuntimeError(f'Unknown input data type '
                           f'({params["input_data_type"]}) provided, only '
                           f'recognized types are (csv, xlsx)')
    if len(data_rows_in) == 0:
        raise RuntimeError(f'Failed to read any data rows from input data '
                           f'file ({params["input_file_path"]}')

    valid_indexes, err_indexes, output_indexes = _calculate_output_indexes(
        in_headers=copy.copy(data_rows_in[0]),
        params=params)
    if not valid_indexes:
        raise RuntimeError(f'failed to translate data_columns to data '
                           f'index numbers: {err_indexes}')

    try:
        with open(params['output_file_path'],
                  'w',
                  encoding=params['character_encoding'],
                  errors=params['character_encoding_errors']) as out_fh:
            csv_writer = csv.writer(out_fh,
                                    delimiter=',',
                                    quotechar='"')
            for row in data_rows_in:
                data_row_out = []
                for index in output_indexes:
                    data_row_out.append(row[index])
                csv_writer.writerow(data_row_out)
    except IOError as ioe:
        raise RuntimeError(f'failed to write lines to OUT_PATH '
                           f'({params["output_file_path"]}: {ioe}')


if __name__ == '__main__':
    try:
        script_dir = os.path.dirname(__file__)
        config_file_path = os.path.join(script_dir, 'csv_shuffle.ini')
        config_read, err_str, config = _read_config(config_file_path)
        if not config_read:
            raise RuntimeError(err_str)
        print(config)
        main(config)
    except RuntimeError as rte:
        print(f'shuffle failed: {rte}')
        sys.exit(1)
