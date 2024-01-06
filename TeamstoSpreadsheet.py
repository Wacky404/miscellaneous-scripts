# script to read csv file and translate it to spreadsheet, probably excel....

from openpyxl import load_workbook

file = r'C:\Users\Wayne\Documents\WorkdayUpdatesandResources.xlsx'
wb_members = load_workbook(file)
num_to_col = {
    1: 'B',
    2: 'C',
    3: 'D',
    4: 'E',
    5: 'F',
}
sheet_names = [
    'Change Management',
    'Comms and Resources',
    'Institution Toolkits',
    'Problem Management',
    'Release Management',
    'SharePoint Resources',
]

for sheet in sheet_names:
    ws_change_management = wb_members[sheet]
    for cell in ws_change_management['A']:
        cell_data = cell.value
        cell_next = int(cell.row) + 1
        # don't know what field for what ???
        check = int()
        if cell_data is not None:
            print(cell.coordinate, 'is not None.')
            print('Then how many to next None?')
            # this is how many cells are next before there is an empty cell
            until_empty = int(0)
            for row in ws_change_management.iter_rows(min_row=cell_next, max_col=cell.column, values_only=True):
                print(row[0])
                if row[0] is not None:
                    until_empty += 1
                else:
                    break
            print('This is how many rows are before the next empty row: ', until_empty)
            i = 1
            while i <= until_empty:
                # figuring out how to manipulate data with until_empty
                row_number = int(cell.row) + i
                # moving data from one column and putting it into B and removing the data from A
                data = ws_change_management['A' + str(row_number)].value
                column = str(num_to_col[i])
                print('This is the column to use: ', column)
                for box in ws_change_management[column]:
                    if column == 'B':
                        check = box.row
                    if box.value is None and column == 'B':
                        ws_change_management[str(num_to_col[i]) + str(box.row)].value = data
                        ws_change_management['A' + str(row_number)].value = None
                        wb_members.save(file)
                        break
                    elif box.row == check and box.value is None:
                        ws_change_management[str(num_to_col[i]) + str(box.row)].value = data
                        ws_change_management['A' + str(row_number)].value = None
                        wb_members.save(file)
                        break
                i += 1
    # moving the name up in the field
    for box in ws_change_management['A']:
        if box.value is None:
            for row in ws_change_management.iter_rows(min_row=box.row, max_col=box.column):
                if row[0].value is not None:
                    next_name_found = ws_change_management['A' + str(row[0].row)].value
                    ws_change_management['A' + str(box.row)].value = next_name_found
                    ws_change_management['A' + str(row[0].row)].value = None
                    wb_members.save(file)
                    break
