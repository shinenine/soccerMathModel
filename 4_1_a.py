import xlrd

excel_path = 'F:\\study\\数学建模\\2021\\tongjimcm2021D体能数据.xlsx'
data = xlrd.open_workbook(excel_path)
# print(4.12*0.3)
#
# print(data.sheet_names())
#
# if data.sheets()[6].cell(3, 1).value is None:
print(data.sheets()[6].cell(3, 1).value)
# teams = []
# for i in range(len(data.sheet_names())):
#     table = data.sheets()[i]
#     rows = table.nrows
#     cols = table.ncols
#     team = []
#     for row in range(rows):
#         print(row)
#         score = table.cell(row, 1).value * 0.15 + table.cell(row, 2).value * 0.15 + table.cell(row, 3).value * 0.1 + \
#                 table.cell(row, 4).value * 0.1 + table.cell(row, 5).value * 0.1 + table.cell(row, 6).value * 0.4
#         team.append(score)
#     teams.append(team)
# print(teams)
