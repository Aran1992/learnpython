from openpyxl import load_workbook

# 要求平均值的列的列表
avg_column_index_list = [2, 3, 4, 6]

# 读取Workbook
workbook = load_workbook('sample.xlsx')
# 要操作的sheet
sheet = workbook.active
# 遍历的行列表
row_list = sheet.iter_rows(values_only=True)
# 下一行是否是数据行标记
next_row_is_data_row = False

sum_list = []
for x in avg_column_index_list:
    sum_list.append(0)
count = 0

print_count = 0

for i, row in enumerate(row_list):
    # 判断这行是否是数据行
    if next_row_is_data_row:
        count = count + 1
        for j, column in enumerate(avg_column_index_list):
            sum_list[j] = sum_list[j] + float(row[column])
    # 判断下一行是否是数据行
    if row[0] == '#':
        next_row_is_data_row = True
    else:
        next_row_is_data_row = False
    if row[0] == '[root]' and count != 0:
        for j, s in enumerate(sum_list):
            avg = s / count
            print('第' + str(avg_column_index_list[j]) + '列的平均值是' + str(avg))
            sum_list[j] = 0
        count = 0
        print_count = print_count + 1
        print('print_count', print_count)

for j, s in enumerate(sum_list):
    avg = s / count
    print('第' + str(avg_column_index_list[j]) + '列的平均值是' + str(avg))
    sum_list[j] = 0
count = 0
print_count = print_count + 1
print('print_count', print_count)
