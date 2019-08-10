# database2excel
数据库表导出excel表格脚本

# 环境
数据库管理系统：MySQL  
Python版本：python 3.x

# 安装程序运行所需要的包
因为使用的包只有两个，懒得生成requirements.txt文件，所以请大家自行使用pip安装  
使用以下两个包即可：  
pymysql  
xlsxwriter  
安装以上两个包  
pip3 install pymysql  
pip3 install xlsxwriter

# 运行脚本
通过database2excel(ip, user, psw, database, table_name)该函数直接执行即可

# 参数说明
ip:数据库管理系统所在的ip地址  
user:数据库管理系统的用户名  
psw:数据库管理系统的密码  
database:数据库管理系统中的某个数据库名称  
table_name:某个数据库中的表的名称

