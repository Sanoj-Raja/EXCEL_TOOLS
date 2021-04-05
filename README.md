Simple Python package to automate the excel based tasks.

Steps to use package.

Wrong Phone Number Finder. This will take the phone number column.
# from excel_tools.excelTools import WrongPhoneNumberFinder
# nameOfExcelFile = 'Filtered_Modified_Excel_Of_Part_1_Fresher_Data_First_3k_Raw_Data.xlsx'
# phoneNumberColumn = 11
# WrongPhoneNumberFinder.findWrongPhoneNumberRow(nameOfExcelFile, phoneNumberColumn)

Filter Down the excel by removing all the rows which doesnot matches the given string list. 
This will take a column for finding the given data.
# from excel_tools.excelTools import ExcelFilter
# nameOfExcelFile = 'Modified_Excel_Of_Part_1_Fresher_Data_First_3k_Raw_Data.xlsx'
# nameOfFilteredExcel = f'Filtered_{nameOfExcelFile}'
# stringsList = ['Delhi/NCR']
# filterByColumn = 14
# ExcelFilter.filterRows(nameOfExcelFile, nameOfFilteredExcel, stringsList, 2, filterByColumn)

Unique string finder in rows of a given column.
# from excel_tools.excelTools import ExcelFilter
# nameOfExcelFile = 'Modified_Excel_Of_Part_1_Fresher_Data_First_3k_Raw_Data.xlsx'
# ExcelFilter.getUniqueValueListOfColumn(nameOfExcelFile, 14)

Report Provider for failed and sucessfull outcomes from an excel.
# from excel_tools.excelTools import Report
# Report.getSummary('Report4.xlsx', 3)

By writing the above code you can start using this package.