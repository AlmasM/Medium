import os
from excel_file import ExcelBook
import pandas as pd

def main():

    cwd = os.getcwd()
    fileName = "SampleSpreadsheet.xlsx"
    
    filePath = os.path.join(cwd, fileName)
    
    excelObject = ExcelBook(filePath)
    workSheetName = "Sheet-List"
    currentWorkSheet = excelObject.getFile(workSheetName)

    if not currentWorkSheet:
        isFileCreated = excelObject.createFile(workSheetName)
    
    data = [ ['Technology Type', 'Market Share(%)'], ['Phone', 50], ['PC', 20], ['Tablets', 20],['Smart Watch', 10]]

    df = pd.DataFrame(data)
    excelObject.addDataFromDataFrame(df)
    df_result = excelObject.getDataAsDataFrame()

    workSheetName = "Sheet-Pandas"
    currentWorkSheet = excelObject.getFile(workSheetName)
    data = [ ['Technology Type', 'Market Share(%)'], ['Phone', 50], ['PC', 20], ['Tablets', 20],['Smart Watch', 10]]
    excelObject.addData(data)

    data_result = excelObject.getDataRowByRow()
    excelObject.createPieChart(data)


if __name__ == '__main__':
    main()