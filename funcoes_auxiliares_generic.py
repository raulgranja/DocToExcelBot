import re
from pathlib import Path
from docx2txt import process
from openpyxl import load_workbook


def date_converter(date):
    """
    Converts a date in the format 'dd month yyyy' or 'dd-month-yyyy' to the format 'dd/mm/yyyy'
    :param date: Date to convert
    :return: Converted date or the original date if the conversion is not possible
    """
    dict_months = {'jan': '01', 'feb': '02', 'mar': '03', 'apr': '04', 
                   'may': '05', 'jun': '06', 'jul': '07', 'aug': '08', 
                   'sep': '09', 'oct': '10', 'nov': '11', 'dec': '12',
                   'january': '01', 'february': '02', 'march': '03',
                   'april': '04', 'june': '06', 'july': '07',
                   'august':'08', 'september': '09', 'october': '10',
                   'november': '11', 'december': '12'}
    try:
        date_list = re.split('-| ', date.lower())
        converted_date = f'{date_list[0]}/{dict_months[date_list[1]]}/{date_list[2]}'
        return converted_date
    
    except:
        return date
    

def load_document(folder_path, doc_name, typefile):
    """
    Loads a document of type docx or xlsx from the specified path
    :param folder_path: Path of the folder where the file is located
    :param doc_name: Name of the file
    :param typefile: file type ('docx' or 'xlsx')
    :return: Contents of the file
    """
    if typefile == 'docx':
        path = folder_path / Path(doc_name)
        return process(path)
    elif typefile == 'xlsx':
        path = folder_path / Path(doc_name + '.xlsx')
        return load_workbook(path)
    
    
def get_string_from_doc(cn_filename, folder):
    """
    Gets a string from the contents of a docx file
    :param cn_filename: File name
    :return: String of file contents
    """
    processed = load_document(folder, cn_filename, 'docx')
    string = re.sub('\n+', ' ', processed)
    return string


def find_between(string, start, end, group):
    """
    This function fetches and returns the value of a specific group between two
    delimiters in a string.
    
    Arguments:
    - string (str): the string on which the search will be performed
    - start (str): the starting delimiter of the group
    - end(str): the end delimiter of the group
    - group(int): the index of the group to be returned (starting with 1)
    
    Return:
    - The group value found, removing leading and trailing spaces.
    """
    itemRegex = re.compile(fr'({start})(.*?)({end})')
    mo = itemRegex.search(string)
    return mo.group(group).strip()


def iter_cells(rows=1748, columns=394):
    """
    This function iterates over the cells of a worksheet, returning the coordinate
    (row, column) of each cell.
    
    Arguments:
    - rows (int): the number of rows in the worksheet (default: 1748)
    - columns (int): the number of columns in the worksheet (default: 394)
    
    Return:
    - A generator with tuples (row, column) representing the coordinates of each
    cell on the worksheet.
    """
    for row in range(2, rows + 2):
        for column in range(2, columns + 2):
            yield row, column
            
            
def get_parameters(params_file, line):
    """
    This function reads a parameter file and returns a specific line.
    
    Arguments:
    - params_file (str): parameter file path
    - line (int): number of the line to be returned
    
    Return:
    - a line from the parameters file
    """
    with open(params_file) as parameters:
        content = parameters.readlines()
        return content[line]
    
    
def generate_first_ID(params_file):
    """
    This function generates the first ID from a parameter file.
    
    Arguments:
    - params_file (str): parameter file path
    
    Return:
    - The first ID generated
    """
    cn_id = get_parameters(params_file, 2)
    return cn_id


def regex_loop(pattern_list, end_counter, cn_subject):
    """
    This function performs a search for a list of patterns in a string,
    and returns the first group found.
    
    Arguments:
    - pattern_list(list): list of regular expression patterns to look for in the string
    - end_counter(int): maximum number of attempts before returning 'nan'
    - cn_subject (str): string in which the search will be performed
    
    Return:
    - The first group found in the string, removing leading spaces and
    finals. If no group is found, returns 'nan'
    """
    counter = 0
    
    for pattern in pattern_list:
        itemRegex = re.compile(pattern)
        mo = itemRegex.search(cn_subject)
        counter += 1
        
        if mo is not None:
            return mo.group().strip()
        
        else:
            if counter == end_counter:
                return 'nan'
            
            
def get_cn_number(cn_subject, number_patterns):
    """
    This function fetches and returns a change number from a string.
    
    Arguments:
    - cn_subject (str): the string on which the search will be performed
    
    Return:
    - The change number found in the string, or 'nan' if not found.
    """
    return regex_loop(number_patterns, 6, cn_subject)


def get_division(cn_subject, number_patterns, stage_patterns):
    """
    This function fetches and returns the division of a change from a string.
    
    Arguments:
    - cn_subject (str): the string on which the search will be performed
    
    Return:
    - The division of the change found in the string, or 'nan' if not found.
    """
    group = 2
    for number_pattern in number_patterns:
        for stage in stage_patterns:
            try:
                division = find_between(cn_subject, number_pattern, stage, group)
                division = division.strip()
                return division
                
            except AttributeError:
                continue
            
            
def get_reference(cn_string, Nonummy=False):
    """
    This function fetches and returns a change reference from a string.
    
    Arguments:
    - cn_string (str): the string on which the search will be performed
    - Nonummy (bool): if True, searches for a change reference in the Nonummy format. If False, search in the format of other BUs (default: False)
    
    Return:
    - The change reference found in the string, 'ERROR' if not found.
    """
    try:
        if Nonummy is False:
            start = 'Massa'
            end = 'Pellentesque'
            reference = find_between(cn_string, start, end, 2)
            return reference
        else:
            start = 'Habitant'
            end = 'Morbi'
            reference = find_between(cn_string, start, end, 2)
            return reference
        
    except:
        return 'ERROR'
    
    
def get_stage(cn_subject, stage_patterns):
    """
    This function fetches and returns the step of a change from a string.
    
    Arguments:
    - cn_subject (str): the string on which the search will be performed
    
    Return:
    - The change step found in the string, or 'nan' if not found.
    """
    return regex_loop(stage_patterns, 3, cn_subject)


def get_implementation_date(cn_string, Nonummy=False):
    """
    This function searches and returns the implementation date of a change from a string.
    
    Arguments:
    - cn_string (str): the string on which the search will be performed
    - Nonummy (bool): if True, searches for the implementation date of a change in the Nonummy format.
      If False, search in standard format (default: False)
    
    Return:
    - The implementation date of the change found in the string,
      and a boolean indicating whether the date found is valid. 'ERROR' if unable to find.
    """
    pattern1 = '\d\d\/\d\d\/\d\d\d\d'
    pattern2 = '\d\d\/\d\d\/\d\d'
    pattern3 = '\d\/\d\d\/\d\d\d\d'
    
    try:
        if Nonummy is False:
            start = 'Tristique'
            end = 'Senectus'
            implementation_date = find_between(cn_string, start, end, 2)
            
        else:
            start = 'Et'
            end = 'Netus'
            implementation_date = find_between(cn_string[329:], start, end, 2)
    
        date = date_converter(implementation_date)
        matched = re.match(fr'{pattern1}|{pattern2}|{pattern3}', date)
        is_match = bool(matched)
        
        return date[:6] + '20' + date[6:], is_match
    
    except:
        return 'ERROR', 'ERROR'


def get_receipt_date(cn_subject):
    """
    This function searches and returns the date of receipt of a change from a string.
    
    Arguments:
    - cn_subject (str): the string on which the search will be performed
    
    Return:
    - The date of receipt of the change found in the string, or 'nan' if not found.
    """
    date = date_converter(cn_subject[-14:-5].strip())
    return date[:6] + '20' + date[6:]  


def get_product(cn_string, Nonummy=False):
    """
    This function fetches and returns the product name of a change from a string.
    
    Arguments:
    - cn_string (str): the string on which the search will be performed
    - Nonummy (bool): if True, searches for the product name of a change in Nonummy format. If False, search in standard format (default: False)
    
    Return:
    - The product name found in the string, or 'ERROR' if not found.
    """
    try:
        if Nonummy is False:
            start = 'Malesuada'
            end = 'Fames'
            product = find_between(cn_string, start, end, 2)
            return product
        else:
            start = 'Ac'
            end = 'Turpis'
            product = find_between(cn_string[329:], start, end, 2)
            return product
        
    except:
        return 'ERROR'
    
    
def get_description(cn_string, Nonummy=False):
    """
    This function fetches and returns the description of a change from a string.
    
    Arguments:
    - cn_string (str): the string on which the search will be performed
    - Nonummy (bool): if True, searches the description of a change in Nonummy format. If False, search in DCAF format (default: False)
    
    Return:
    - The description of the change found in the string, or 'ERROR' if not found.
    """
    try:
        if Nonummy is False:
            start = 'Egestas'
            end = 'Nulla'
            description = find_between(cn_string, start, end, 2)
    
        else:
            start = 'Risus'
            end = 'Quisque'
            description = find_between(cn_string[:], start, end, 2)

        return description
    
    except:
        return 'ERROR'


def find_x(string):
    """
    This function searches for "X NO", "X YES" or "X N/A" in a string and returns "No", 
    "Yes" or "N/A", respectively.
    
    Arguments:
    - string (str): the string on which the search will be performed
    
    Return:
    - "No" if the string contains "X NO ", "Yes" if it contains "X YES " or "N/A" if it contains "X N/A",
    or "ERROR" if none of these cases are found.
    """
    string = string.upper()
    if string.find('X NO ') != -1:
        return 'No'
    elif string.find('X YES ') != -1:
        return 'Yes'
    elif string.find('X N/A') != -1:
        return 'N/A'
    else:
        return 'ERROR'


def string_iterator(string):
    """
    This function iterates over a string and adds '\' before special characters.
    
    Arguments:
    - string (str): the string on which the operation will be performed
    
    Return:
    - string with special characters preceded by '\'
    """
    count = 0
    for i, character in enumerate(string):
        if character in ('/.[]{}()<>*+-=!?^$|'):
            string = string[:i + count] + '\\' + string[i + count:]
            count += 1
    string_processed = string
    
    return string_processed
