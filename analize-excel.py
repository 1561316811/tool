import xlrd
import xlwt

data = xlrd.open_workbook('/workspace/ProtocolFileCovert2LinkDoc/data/2022-5-27号物质发放登记.xls') # 打开xls文件

sheet = data.sheets()[0] # 打开第一张表
nrows = sheet.nrows # 获取表的行数
columns=sheet.ncols 

print(nrows,columns)

room_stu = {}
room_set = set()
names = []
for i in range(nrows):
    if i == 0: # 跳过第一行
       continue
    room = ''
    for j in range(columns):
        if j == 0: # 统计房间号码
            room = sheet.cell(i,j)
            if room.ctype == 2:
                room = str(int(room.value))
            elif room.ctype == 0:
                continue
            else:
                room = str(room.value)
            if not room_stu.keys().__contains__(room):
                room_stu[room] = []
        else: #统计人数
            stu = sheet.cell(i,j)
            if stu.ctype == 0:
                continue
            room_stu[room].append(stu.value)


test_data_1=sorted(room_stu.items(),key=lambda x:x[0]) 
print(test_data_1)

# 创建一个Workbook对象 编码encoding
Excel = xlwt.Workbook(encoding='utf-8', style_compression=0)

# 添加一个sheet工作表、sheet名命名为Sheet1、cell_overwrite_ok=True允许覆盖写
table = Excel.add_sheet('Sheet1', cell_overwrite_ok=True)
for row in range(len(test_data_1)):
    for col in range(len(test_data_1[row])):
        if col == 0:
            table.write(row, col, test_data_1[row][col])
        else:
            k = 0
            for stu in test_data_1[row][col]:
                table.write(row, col+k, test_data_1[row][col][k])
                k += 1

Excel.save(r'./2022-5-28物质领取登记.xls')
