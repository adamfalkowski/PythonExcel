from openpyxl import Workbook, load_workbook
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.utils import * 
import os.path




def transfer(operation, original_workbook_name, new_workbook, new_worksheet_name, spreadsheet_name):
    
    original_workbook = load_workbook(original_workbook_name)
    original_workbook_sheet = original_workbook.active

    new_workbook_sheet = new_workbook.active

    if operation == "defined names":

        orignal_def_names = original_workbook.defined_names.definedName #List of defined names

        for name in orignal_def_names:
            
            if name.value != "#REF!":
                _ , second_value = name.value.split("!")
                new_def_names = DefinedName(name=name.name,attr_text=f"{new_worksheet_name}!{second_value}")
                new_workbook.defined_names.append(new_def_names)
        
        new_workbook.save(spreadsheet_name) 
    

    elif operation == "column headers":

        for cell in original_workbook_sheet[1]: 
            new_workbook_sheet[cell.coordinate] = cell.value
            

        new_workbook.save(spreadsheet_name) 

    else:
     return 0







def main():

    # Project variables
    spreadsheet_name = "TagsEquipment_6200000120.xlsx"
    folder_path = "D:\Practice\Python"
    file_path = folder_path + spreadsheet_name


    if not os.path.exists(file_path):
        tags_equipment = Workbook()
        tags_equipment_sheet = tags_equipment.active
        tags_equipment_sheet.title = "Tags_Equipment"
        tags_equipment.save(spreadsheet_name) 
    else:
        tags_equipment = load_workbook(file_path)
        

    
    transfer("defined names","TagsEquipment_bgfcv_templete.xlsx",tags_equipment, "Tags_Equipment", spreadsheet_name)
    transfer("column headers","TagsEquipment_bgfcv_templete.xlsx",tags_equipment, "Tags_Equipment", spreadsheet_name)
    #transfer_column_names("TagsEquipment_bgfcv_templete.xlsx",tags_equipment, "Tags_Equipment",spreadsheet_name)


main()
    
