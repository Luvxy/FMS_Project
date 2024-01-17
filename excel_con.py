import pandas as pd
from tkinter import filedialog as fd

path = ''

filetypes = (
    ('excel files', '*.xlsx'),
    ('All files', '*.*')
)

filename = fd.askopenfilename(
    title='Open a file',
    initialdir='/',
    filetypes=filetypes)

path = filename
df = pd.read_excel(path)

# Find a specific value in the Excel sheet
def find_value(value):
    result = df[df == value].stack()
    if not result.empty:
        return result.index[0]

# Edit a cell value
def edit_cell(row, column, new_value):
    df.at[row, column] = new_value

# Delete a row
def delete_row(row):
    df.drop(row, inplace=True)

# Modify a cell value solve error out of range
def modify_cell(row, column, modification):
    try:
        current_value = df.at[row, column]
        modified_value = modification(current_value)
        df.at[row, column] = modified_value
    except:
        print('IndexError: row or column out of range')
    
# Find a specific value in the Excel sheet and modify it
def find_and_modify_value(value, modification):
    result = df[df == value].stack()
    if not result.empty:
        row = result.index[0][0]
        column = result.index[0][1]
        current_value = df.at[row, column]
        modified_value = modification(current_value)
        df.at[row, column] = modified_value
        
# Main
if __name__ == '__main__':
    # Find a specific value in the Excel sheet
    print(find_value('value'))
    
    # Edit a cell value
    edit_cell(0, 'column', 'new_value')
    
    # Delete a row
    delete_row(0)
    
    # Modify a cell value
    modify_cell(0, 'column', lambda x: x + 1)
    
    # Find a specific value in the Excel sheet and modify it
    find_and_modify_value('value', lambda x: x + 1)

    # Save the changes to the Excel file
    df.to_excel('C:/Users/user/Desktop/test/test_excel.xlsx', index=False)

