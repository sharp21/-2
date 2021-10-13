# from service.ITestCaseRead import ITestCaseRead
# import openpyxl
# from model.TestModel import TestModel
# from model.TestModel import TestData
#
#
# class TestCaseReadFromExcel(ITestCaseRead):
#     def __init__(self, fileName):
#         self.fileName = fileName
#
#     def readCell(self, sheet, row, column, defaultValue):
#         value = sheet.cell(row=row, column=column).value
#         if value is None:
#             value = defaultValue
#         return value
#
#     def readTestCase(self):
#         wb = openpyxl.load_workbook(self.fileName)
#         sheets = wb.sheetnames
#         models = []
#         for i in sheets:
#             sheet = wb.get_sheet_by_name(i)
#             model = TestModel(sheet.title)
#             for row in range(2, sheet.max_row + 1):
#                 name = self.readCell(sheet, row, 1, '-')
#                 url = self.readCell(sheet, row, 2, '')
#                 method = self.readCell(sheet, row, 3, '')
#                 inputParams = self.readCell(sheet, row, 4, '')
#                 checkMode = self.readCell(sheet, row, 5, '')
#                 checkUrl = self.readCell(sheet, row, 6, '')
#                 checkMethod = self.readCell(sheet, row, 7, '')
#                 checkInputParams = self.readCell(sheet, row, 8, '')
#                 checkPoint = self.readCell(sheet, row, 9, '')
#                 data = TestData(name, url, method, inputParams)
#                 data.setCheckParam(checkMode, checkUrl, checkMethod, checkInputParams, checkPoint)
#                 model.addTestData(data)
#             models.append(model)
#         return models
