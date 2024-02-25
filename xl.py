from openpyxl import Workbook, load_workbook
from openpyxl.workbook.defined_name import DefinedName

# Project variables
project_number = "6200000120"




def transfer_defined_names(original_workbook_name, new_workbook_name, new_worksheet_name, project_number):
    original_workbook = load_workbook(original_workbook_name)
    orignal_def_names = original_workbook.defined_names.definedName

    new_workbook = Workbook()
    new_workbook_sheet = new_workbook.active
    new_workbook_sheet.title = new_worksheet_name

    for name in orignal_def_names:
        
        if name.value != "#REF!":
            first_value, second_value = name.value.split("!")
            new_def_names = DefinedName(name=name.name,attr_text=f"{new_worksheet_name}!{second_value}")
            new_workbook.defined_names.append(new_def_names)

    new_workbook.save(f"{new_workbook_name}_{project_number}.xlsx")

    return new_workbook

transfer_defined_names("TagsEquipment_bgfcv_templete.xlsx","TagsEquipment", "Tags_Equipment",project_number)



