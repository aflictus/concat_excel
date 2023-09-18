import pandas as pd
import os

FILE_NAME = 'Example.xlsx'
EXCEL_PATH = os.getcwd() # PWD
file_path = "/".join([EXCEL_PATH, FILE_NAME])

class Excel:
    def __init__(self, sheet_name: str, data: [list or dict]):
        self.file = pd.read_excel(file_path, sheet_name=sheet_name)
        self.sheet_name = sheet_name   
        self.data = pd.DataFrame(data)                        
        self.new_data = pd.concat([self.file, self.data], ignore_index=True)
        self.new_data = self.new_data.fillna(pd.NaT) # Изменение значения в скобках приведёт к изменению значениЙ во всей таблице


class ExcelWriter:
    def __init__(self, sheets: list):
        self.sheets = sheets

    def write_data(self):
        with pd.ExcelWriter(file_path) as writer:
            for sheet in self.sheets:
                sheet.new_data.to_excel(writer, sheet_name=sheet.sheet_name, index=False)


""" ПРИМЕР """
'''
# Пример того как выглядит дата
collected_data = [{0: 1, 1: 2, 2: 3, 3: 4, 4: None, 5: None, 6: None, 7: 8, 8: 9}, {0: 1, 1: 2, 2: 3, 3: 4, 4: None, 5: None, 6: None, 7: 8, 8: 9}, {0: 1, 1: 2, 2: 3, 3: 4, 4: None, 5: None, 6: None, 7: 8, 8: 9}]
collected_data2 = [[1, 2, 3, 4, None, None, None, 8, 9], [1, 2, 3, 4, None, None, None, 8, 9], [1, 2, 3, 4, None, None, None, 8, 9]]

# Пример того как создавать объекты класса
example = Excel(sheet_name='Example', data=collected_data)
sheet2 = Excel(sheet_name='Sheet2', data=collected_data2)

# Пример того как добавлять новые данные
ew = ExcelWriter([example, sheet2])
ew.write_data()  
'''
