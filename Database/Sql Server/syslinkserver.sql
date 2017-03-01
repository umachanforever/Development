  
--openquery 用法
---https://msdn.microsoft.com/zh-cn/library/ms188427.aspx
--查询当前连接的数据库
select * from sys.servers
--查询远程数据库数据
--示例 ,依次为sql server和oracle
select  top 10 * from [192.168.100.100].test.dbo.[demo_table]
    
select * from openquery(oracledb,'select * from oracletest.demo_table where id=11111')

--sp_addlinkedserver 用法
--http://www.cnblogs.com/nov5026/p/6052119.html
--https://msdn.microsoft.com/en-us/library/ms190479.aspx
--http://blog.sina.com.cn/s/blog_553852910100rmsi.html

