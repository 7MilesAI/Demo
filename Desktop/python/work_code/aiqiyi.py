#!/usr/bin/env python
# -*- coding:utf-8 -*-
# Author:Tom Hardy

import cx_Oracle
import pandas as pd
from datetime import datetime, timedelta
import pymysql

# ========================================================

conn = cx_Oracle.connect('bonan/bonan@172.25.88.52:1521/jldata')

cursor = conn.cursor()

sql_tjw = """
SELECT
	to_char( time, 'yyyy-mm-dd' ) as "日期",
	count(userid) as "访问次数",
	count(distinct userid) as "访问人数"
FROM
	dx_column_user 
WHERE
	columnid = '20001110000000000000000000004600'
	and time >= to_date( '20200117', 'yyyy-mm-dd' ) 
group by 
	to_char( time, 'yyyy-mm-dd' )
ORDER BY
	to_char( time, 'yyyy-mm-dd' )

"""

cursor.execute(sql_tjw)

data_tjw = cursor.fetchall()

result_tjw = pd.DataFrame(data=data_tjw, columns=["日期", "访问次数", "访问人数"])

print("ok!")

# ====================================================================================================
conn_mysql = pymysql.connect(host='10.128.7.7',
                             user='queryuser',
                             passwd='1qaz@WSX',
                             db='jlct_vas',
                             port=3306,
                             charset='utf8')

cursor_dx = conn_mysql.cursor()

sql_dj = """
select
substr(paytime, 1, 10) as "日期",
sum(fee)/100 "订购金额",
count(userid) as "订购量"
from
tab_order
where productid = '1100000181'
and status = 1
and substr(paytime, 1, 10) >= '2020-01-01'
group by substr(paytime, 1, 10)
order by substr(paytime, 1, 10) 

"""

cursor_dx.execute(sql_dj)

data_dj = cursor_dx.fetchall()

result_dj = pd.DataFrame(data=data_dj, columns=["日期", "订购金额", "订购量"])

print("ok!")

# =======================================================================================================


sql_user = """
select
a1.userid,
case
	when a2.userid is not null and a3.userid is not null then '连续两个月订购任意产品'
	end as "是否连续订购",
case
    when a4.userid is null then '近半年无订购'
    end as  "近半年是否订购"
from
(SELECT
	distinct userid
FROM
	dx_column_user 
WHERE
	columnid = '20001110000000000000000000004600'
	and time = trunc(sysdate) - 1
-- time = to_date('2020-01-25', 'yyyy-mm-dd')
	and userid not in (
select 
	distinct userid 
from
	tab_order
where productid = '1100000181'
	and status = 1
	and substr(paytime, 1, 10) >= '2020-01-01'
	)) a1
left join 
(select 
	distinct userid 
from
	tab_order
where
	status = 1
and substr(paytime, 1, 10) >= '2019-12-01'
and substr(paytime, 1, 10) < '2020-01-01')a2
on a1.userid = a2.userid
left join 
(select 
	distinct userid 
from
	tab_order
where status = 1
and substr(paytime, 1, 10) >=  '2019-11-01'
and substr(paytime, 1, 10) <  '2019-12-01')a3
on a1.userid = a3.userid
left join 
(select 
	distinct userid 
from
	tab_order
where status = 1
and substr(paytime, 1, 10) >=  '2019-06-01'	)a4
on a1.userid = a4.userid

"""

cursor.execute(sql_user)

data_user = cursor.fetchall()

result_user = pd.DataFrame(data=data_user, columns=["IPTV账号", "是否连续订购", "近半年是否订购"])

print("ok!")

# ========================================================================================================
# 前一天日期
date_rand = (datetime.now() + timedelta(days=-1)).strftime('%Y-%m-%d')


path_data = r'C:\Users\yangshen\Desktop\工作内容\2_task\爱奇艺数据\\'
file_name_data = path_data + '%s爱奇艺数据.xlsx' % date_rand
writer_data = pd.ExcelWriter(file_name_data)

# 写入缓存
result_tjw.to_excel(writer_data, sheet_name="推荐位", index=0)
result_dj.to_excel(writer_data, sheet_name="订购", index=0)


path_user = r'C:\Users\yangshen\Desktop\工作内容\2_task\爱奇艺用户\\'
file_name_user = path_user + '%s爱奇艺用户.xlsx' % date_rand
writer_user = pd.ExcelWriter(file_name_user)

result_user.to_excel(writer_user, sheet_name="浏览没订购爱奇艺", index=0)

writer_data.save()
writer_data.close()

writer_user.save()
writer_user.close()

cursor.close()
conn.close()
