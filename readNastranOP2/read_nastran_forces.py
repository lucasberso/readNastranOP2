
from pyNastran.op2.op2 import OP2
import xlwings as xw
import pandas as pd

# PATHS FOR THE INPUT OP2 AND EXCEL BOOK:
file = 'C:\\Users\\lucas.bernacer\\Documents\\Herramienta Nastran\\A320_ESGPlus_S18_ALL.op2'
excel_book = 'C:\\Users\\lucas.bernacer\\Documents\\Herramienta Nastran\\Excel_Nastran_forces_template.xlsm'
sheetname = "Read NASTRAN GFEM"

# FUNCTIONS:
def read_input_excel(excel_book_path, sheetname):
    """
    Reads the information regarding the Nastran subcases and elements' ids from the Excel input file. The Excel sheet
    must be named "Read_Nastran_force".
    Nastran's subcases definition range: From cells A12:A811.
    Nastran's elements definition range: D6:W6.

    :param excel_book_path: Path of the input Excel book.
    :return: Returns the subcases list and a dictionary with the element information.
    """
    book = xw.Book(excel_book_path)
    sheet = book.sheets[sheetname]
    element_dict = {}
    subcase_list = [int(i) for i in sheet.range("A12:A811").value if i] # Stores the subcases information.
    count = 0
    for col in range(4, 116): # Gathers the elements' dictionary.
        row = 9
        if sheet.range(row, col).value:
            element = int(sheet.range(row, col).value)
            element_type = sheet.range(row+1, col).value
            element_force_type = sheet.range(row+2, col).value
            if element_type and element_force_type:
                element_dict.update({count:[element,element_type,element_force_type]})
        count = count + 1
    return subcase_list, element_dict

def write_output_excel(excel_book_path, output_dict, sheetname):
    """
    Writes the output force information for each element extracted from the Excel input file.
    Starting writing from the Excel input file: D12. The Excel sheet must be named "Read_Nastran_force".

    :param excel_book_path: Path of the input Excel book.
    :param output_dict: Information containing the force information per element.
    """
    book = xw.Book(excel_book_path)
    sheet = book.sheets[sheetname]
    sheet.range("D12").options(index=False, header=False).value = pd.DataFrame(output_dict)

def check_subcase_element(element_id, element_type, subcase_id, forces_dict):
    """
    Verifies that the element force information is stored in the OP2 input file.

    :param element_id: Nastran's element id.
    :param element_type: Type of Nastran element between CBAR, CROD and CQUAD.
    :param subcase_id: Nastran's subcase.
    :param forces_dict: OP2's dictionary with the force information.
    :return: Returns a flag that verifies if the element force information is stored in the OP2 file and
    additional variables to retrieve it.
    """
    check, force_key, element_index = False, None, None
    force_key_dict = {'CBAR': 'force.cbar_force',
                      'CROD': 'force.crod_force',
                      'CQUAD': 'force.cquad4_force',
                      'CTRIA': 'force.ctria3_force'}

    if element_type in list(force_key_dict.keys()): # Checks if the element type is supported.
        force_key = force_key_dict[element_type]
    if force_key:
        subcase_list = list(forces_dict[force_key].keys())
        if subcase_id in subcase_list: # Checks if the element has been solved for the evaluated subcase.
            if force_key in ['force.cbar_force', 'force.crod_force', 'force.cquad4_force', 'force.ctria3_force']:
                element_list = forces_dict[force_key][subcase_id].element.tolist()
            else:
                element_list = []
            if element_id in element_list:
                check, element_index = True,  element_list.index(element_id)
    return check, force_key, element_index

# CODE:
op2 = OP2()
op2.read_op2(file, True, False, True, None) # Reads the Nastran OP2 file.
forces_dict = {}
for i in list(op2.__dict__.keys()): # Retrieves the force information stored in the OP2 file.
    if 'force' in i:
        if getattr(op2, i):
            forces_dict.update({i:getattr(op2, i)})

subcase_list, element_dict = read_input_excel(excel_book, sheetname) # Function that gets the subcases and elements information.
force_type_dict = {'force.cbar_force':{'MA1':0, 'MA2':1, 'MB1':2, 'MB2':3, 'V1':4, 'V2':5, 'FA':6},
                   'force.crod_force':{'FA':0},
                   'force.cquad4_force':{'FX':0, 'FY':1, 'FXY':2},
                   'force.ctria3_force':{'FX':0, 'FY':1, 'FXY':2}}
output_dict = {}

for i in list(element_dict.keys()): # Loop that extracts the elements' force information and stores it.
    element = element_dict[i][0]
    subcases_force_list = []
    for subcase in subcase_list: # Checks the validity of the input element and subcase.
        check, force_key, index = check_subcase_element(element, element_dict[i][1], subcase, forces_dict)
        if check:
            if element_dict[i][2] in list(force_type_dict[force_key].keys()):
                force_data_index = force_type_dict[force_key][element_dict[i][2]]
                subcases_force_list.append(forces_dict[force_key][subcase].data[0][index][force_data_index])
            else:
                subcases_force_list.append('')
        else:
            subcases_force_list.append('')
    output_dict.update({i:subcases_force_list}) # Appends the force data to a global dictionary of elements.
write_output_excel(excel_book, output_dict, sheetname) # Writes the information in the input Excel book.
