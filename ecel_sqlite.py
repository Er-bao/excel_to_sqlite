#实现excel转数据库文件
'''author:白胖胖
   time:2022.4.6
'''

import sqlite3
import xlrd

class ExcelToSqlite(object):
    exe = "     执行: "
    output = "     输出: "
    sheetDataStartIndex = 1  # 数据开始计算的行数，如第0行是表头，第1行及之后是数据
 
    def __init__(self, dbName):
        print("初始化数据库实例")
        super(ExcelToSqlite, self).__init__()
        self.conn = sqlite3.connect(dbName)
        self.cursor = self.conn.cursor()
 
    def __del__(self):
        print("释放数据库实例")
        self.cursor.close()
        self.conn.close()
 
    def ExcelToDb(self, excelName, sheetIndex, tableName):
        """
        excel转化为sqlite数据库表
        :param excelName:excel名
        :param sheetIndex:excel中sheet位置
        :param tableName:数据库表名
        """
        print("Excel文件 转 db")
        self.tableName = tableName
        excel = xlrd.open_workbook(excelName)
        sheet = excel.sheets()[sheetIndex]  # sheets 索引
        self.sheetRows = sheet.nrows  # excel 行数
        self.sheetCols = sheet.ncols  # excle 列数
        fieldNames = sheet.row_values(0)  # 得到表头字段名
        print('表头字段')
        print(fieldNames)
        # 创建表
        fieldTypes = ""
        for index in range(fieldNames.__len__()):
            if (index != fieldNames.__len__() - 1):
                fieldTypes += fieldNames[index] + " text,"
            else:
                fieldTypes += fieldNames[index] + " text"
        self.__CreateTable(tableName, fieldTypes)
        # 插入数据
        for rowId in range(self.sheetDataStartIndex, self.sheetRows):
            fieldValues = sheet.row_values(rowId)
            self.__Insert(fieldNames, fieldValues)
 
    def __CreateTable(self, tableName, field):
        """
        创建表
        :param tableName: 表名
        :param field: 字段名及类型
        :return:
        """
        print('#########################################')
        print("创建表 " + tableName)
        sql = 'create table if not exists %s(%s)' % (self.tableName, field)  # 这里注意要切片，插入的表头格式有多余逗号出现
        print(self.exe + sql)#执行xx指令
        self.cursor.execute(sql)
        self.conn.commit()

    def __Insert(self, fieldNames, fieldValues):
        """
        插入数据
        :param fieldNames: 字段list
        :param fieldValues: 值list
        """
        # 通过fieldNames解析出字段名
        names = ""  # 字段名，用于插入数据
        nameTypes = ""  # 字段名及字段类型，用于创建表
        for index in range(fieldNames.__len__()):
            if (index != fieldNames.__len__() - 1):
                names += fieldNames[index] +","
                nameTypes += fieldNames[index] + " ,text"
            else:
                names += fieldNames[index]
                nameTypes += fieldNames[index] + " text"
        # 通过fieldValues解析出字段对应的值
        values = ""
        for index in range(fieldValues.__len__()): # 读取的excel取值，注意循环次数，也要调整不然会出现多余的，空值
            cell_value = str((fieldValues[index]))
            if (isinstance(fieldValues[index], float)):
                cell_value = str((int)(fieldValues[index]))  # 读取的excel数据会自动变为浮点型，这里转化为文本
            if (index != fieldValues.__len__() - 1):
                values += "\'" + cell_value + "\',"
            else:
                values += "\'" + cell_value + "\'"
        # 将fieldValues解析出的值插入数据库
        sql = 'insert into %s (%s) values(%s)' % (self.tableName, names, values)  # 读取的excel千万注意要切片，否则会多出没用的符号
        print(self.exe + sql)
        self.cursor.execute(sql)
        self.conn.commit()
 
    def Query(self, tableName):
        """
        查询数据库表中的数据
        :param tableName:表名
        """
        print('#########################################')
        print("查询表 " + tableName)
        sql = 'select * from %s' % (tableName)
        print(self.exe + sql)
        self.cursor.execute(sql)
        results = self.cursor.fetchall()  # 获取所有记录列表
        index = 0
        for row in results:
            print(self.output + "index=" + index.__str__() + " detail=" + str(row))  # 打印结果
            index += 1
        print(self.output + "共计" + results.__len__().__str__() + "条数据")

    def executeSqlCommand(self, sqlCommand):
        """
        执行输入的sql命令
        :param sqlCommand: sql命令
        """
        print("执行自定义sql " + tableName)
        print(self.exe + sqlCommand)
        self.cursor.execute(sqlCommand)
        results = self.cursor.fetchall()
        print(self.output + str(results))
        for index in range(0, results.__len__()):
            print(self.output + str(results[index]))
        self.conn.commit()

#只需要修改如下参数即可导入成功
sqlite3.connect('test0.db') # 自己利用sqlit3创建一个database
dbName = "test0.db"  # 把数据库名赋值给函数变量
tableName = "test7"  # 数据库 表 名，表存不存在都可以，赋值后，代码会自动创建这个表。
excelName = "test.xls"  # excel名(可加路径)


#运行过程
es = ExcelToSqlite(dbName)
es.ExcelToDb(excelName, 0, tableName)
es.Query(tableName)
#es.executeSqlCommand("select * from " + tableName)
es.executeSqlCommand("select 姓名 from  "+tableName+"  where 班级=2" )
