
from pyNastran.op2.op2 import OP2
import xlwings as xw
import pandas as pd

# PATHS FOR THE INPUT OP2 AND EXCEL BOOK:
file = 'C:\\Users\\lucas.bernacer\\Documents\\Herramienta Nastran\\A320_ESGPlus_S18_ALL_Stress.op2'
excel_book = 'C:\\Users\\lucas.bernacer\\Documents\\Herramienta Nastran\\Excel_Nastran_stress_template.xlsm'
sheetname = "Read NASTRAN GFEM"

# FUNCTIONS:
def read_input_excel(excel_book_path, sheetname):
    """
    Reads the information regarding the Nastran subcases and elements' ids from the Excel input file. The Excel sheet
    must be named "Read_Nastran_Stress".
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
            element_stress_type = sheet.range(row+2, col).value
            if element_type and element_stress_type:
                element_dict.update({count:[element, element_type,element_stress_type]})
        count = count + 1
    return subcase_list, element_dict

def write_output_excel(excel_book_path, output_dict, sheetname):
    """
    Writes the output stress information for each element extracted from the Excel input file.
    Starting writing from the Excel input file: D12. The Excel sheet must be named "Read_Nastran_Stress".

    :param excel_book_path: Path of the input Excel book.
    :param output_dict: Information containing the stress information per element.
    """
    book = xw.Book(excel_book_path)
    sheet = book.sheets[sheetname]
    sheet.range("D12").options(index=False, header=False).value = pd.DataFrame(output_dict)

def check_subcase_element(element_id, element_type, subcase_id, stress_dict):
    """
    Verifies that the element stress information is stored in the OP2 input file.

    :param element_id: Nastran's element id.
    :param element_type: Type of Nastran element between CBAR, CROD and CQUAD.
    :param subcase_id: Nastran's subcase.
    :param stress_dict: OP2's dictionary with the stress information.
    :return: Returns a flag that verifies if the element stress information is stored in the OP2 file and
    additional variables to retrieve it.
    """
    check, stress_key, element_index = False, None, None
    stress_key_dict = {'CBAR': 'cbar_stress', 'CROD': 'crod_stress', 'CQUAD': 'cquad4_stress'}

    if element_type in list(stress_key_dict.keys()): # Checks if the element type is supported.
        stress_key = stress_key_dict[element_type]
    if stress_key:
        subcase_list = list(stress_dict[stress_key].keys())
        if subcase_id in subcase_list: # Checks if the element has been solved for the evaluated subcase.
            if stress_key in ['cbar_stress', 'crod_stress']:
                element_list = stress_dict[stress_key][subcase_id].element.tolist()
            elif stress_key == 'cquad4_stress':
                element_list = [item[0] for item in stress_dict['cquad4_stress'][subcase_id].element_node.tolist()]
            else:
                element_list = []
            if element_id in element_list:
                check, element_index = True,  element_list.index(element_id)
    return check, stress_key, element_index

# CODE:
op2 = OP2()
op2.read_op2(file, True, False, True, None) # Reads the Nastran OP2 file.
stress_dict = {}
for i in list(op2.__dict__.keys()): # Retrieves the stress information stored in the OP2 file.
    if 'stress' in i:
        if getattr(op2, i):
            stress_dict.update({i:getattr(op2, i)})

subcase_list, element_dict = read_input_excel(excel_book, sheetname) # Function that gets the subcases and elements information.
stress_type_dict = {'cbar_stress':{'SA':4},
                    'crod_stress':{'SA':0},
                    'cquad4_stress':{'VM':7}}
output_dict = {}

for i in list(element_dict.keys()): # Loop that extracts the elements' stress information and stores it.
    element = element_dict[i][0]
    subcases_stress_list = []
    for subcase in subcase_list: # Checks the validity of the input element and subcase.
        check, stress_key, index = check_subcase_element(element, element_dict[i][1], subcase, stress_dict)
        if check:
            if element_dict[i][2] in list(stress_type_dict[stress_key].keys()):
                stress_data_index = stress_type_dict[stress_key][element_dict[i][2]]
                subcases_stress_list.append(stress_dict[stress_key][subcase].data[0][index][stress_data_index])
            else:
                subcases_stress_list.append('')
        else:
            subcases_stress_list.append('')
    output_dict.update({i:subcases_stress_list}) # Appends the stress data to a global dictionary of elements.
write_output_excel(excel_book, output_dict, sheetname) # Writes the information in the input Excel book.
