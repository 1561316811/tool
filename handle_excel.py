import xlrd

data = xlrd.open_workbook('data.xls') # 打开xls文件
model = xlrd.open_workbook('model.xls') # 打开xls文件

table_model = model.sheets()[0] # 打开第一张表
nrows = table_model.nrows # 获取表的行数

stus = {}
names = []
for i in range(nrows):
   if i == 0: # 跳过第一行
       continue
   name = table_model.row_values(i)[0]
   names.append(name)
   stus[name] = 0
   # print(name)

table = data.sheets()[0] # 打开第一张表
nrows = table.nrows # 获取表的行数


for i in range(nrows):

   if i == 0: # 跳过第一行
       continue
   name = table.row_values(i)[8]
   score = int(str(table.row_values(i)[9]).split(".")[1])
   is_c = True if str(table.row_values(i)[18]).find('A') != -1 else False
   if is_c:
       score += 0.5
   stus[name] = score
   # stus.append(stu(name,score))
   # print(name, is_c, score)
   # print(table.row_values(i)[:13]) # 取前十三列数据

for n in names:
    print(stus[n])