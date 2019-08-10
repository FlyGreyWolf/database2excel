import pymysql
import json
import xlsxwriter
import time


batch_size = 5
excel_cur_row = 1
table_head = []

def database2excel(ip, user, psw, database, table_name):
    start = time.clock()
    # 打开数据库连接
    db = pymysql.connect(ip, user, psw
                         , database,
                         cursorclass=pymysql.cursors.SSCursor, port=3306, charset='utf8')
    # 使用 cursor() 方法创建一个游标对象 cursor
    cursor = db.cursor()
    sql = "select * from " + table_name

    try:
        cursor.execute(sql)

        des = cursor.description  # 显示每列的详细信息
        # print("表的描述:", des)
        len_des = len(des) #字段的个数，即列的个数
        for i in range(len_des):
            table_head.append(des[i][0])
        print("table_head===>>", table_head)

        workbook, worksheet = init_excel(table_head)  # initial excel

        result = cursor.fetchmany(batch_size)
        count = 0
        index = 0

        while result is not None and result != []:
            list = []
            while (count < cursor.rownumber):
                print(result[index])  # 处理每一行
                dict = {}
                for i in range(len_des):
                    dict[table_head[i]] = result[index][i]
                list.append(dict)

                count = count + 1
                index = index + 1
            if (len(list) > 0):
                print("===>>", list)
                generate_excel(worksheet, list, len_des)

            result = cursor.fetchmany(batch_size)
            index = 0
    except Exception as e:
        db.rollback()  # 如果出错就回滚并且抛出错误收集错误信息。
        print("Error!:{0}".format(e))
    finally:
        cursor.close();  # 关闭游标资源
        db.close()  # 关闭数据库连接
        close_workbook(workbook) #关闭workbook资源
        end = time.clock()
        print("总耗时:", end - start, '秒')

#初始化excel表格
def init_excel(table_head):
    workbook = xlsxwriter.Workbook('./rec_data.xlsx')
    worksheet = workbook.add_worksheet()

    # 设定格式，等号左边格式名称自定义，字典中格式为指定选项
    # bold：加粗，num_format:数字格式
    bold_format = workbook.add_format({'bold': True})

    # 将二行二列设置宽度为20(从0开始)
    # worksheet.set_column(2, 2, 20)
    # worksheet.set_column(3, 3, 20)

    # 用符号标记位置，例如：A列1行
    table_head_len = len(table_head)
    _col = 'A'
    for i in range(table_head_len):
        col = _col + '1'
        worksheet.write(col, table_head[i], bold_format)
        _col = chr(ord(_col) + 1)

    return workbook, worksheet

# 生成excel文件
def generate_excel(worksheet, expenses, col_num):
    global excel_cur_row
    global table_head

    col = 0
    for item in (expenses):
        # 使用write_string方法，指定数据格式写入数据
        for i in range(col_num):
            worksheet.write_string(excel_cur_row, col+i, str(item[table_head[i]]))
        excel_cur_row += 1

#关闭workbook资源
def close_workbook(workbook):
    workbook.close()

if __name__ == '__main__':
    ip = "127.0.0.1"
    user = "root"
    psw = "root"
    database = "test_database"
    table_name = "test_table"

    database2excel(ip, user, psw, database, table_name)

