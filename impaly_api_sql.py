# -*- coding: utf-8 -*-
from impala.dbapi import connect
import pymssql
conn = connect(host='10.65.1.101', port=21050)
spamlog = conn.cursor()
sql = """
select message_id, date_time, src_ip, sender, receiver, title, attach_info, flag, delivery_status, filter_info, certify_yn from spam 
where date_time >= date_trunc('hour', hours_add(now(), 8)) and date_time < date_trunc('hour', hours_add(now(), 9))  order by certify_detail desc """
# 스케쥴 (매시 50분에 전 시간 로그 수집), 13:50 의 수집로그는 12:00 ~ 13:00
#flag 0 : 수신 -1 : 송신
#delivery_status -1 거부 : -2 : 수신완료 0 : 전달
#filter_info : 정책 명
#certify_yn 1 : 인증됨 0 : 인증안됨
#certify_detail
spamlog.execute(sql)
platdb = pymssql.connect(host='10.65.0.122:51433', user='mskim2018', password='Rlarla18!@', database='spamlog')
usermaster = platdb.cursor()
insertsql = """insert into tbspam (message_id, date_time, src_ip, sender, receiver, title, attach_info, flag, delivery_status, filter_info, certify_yn)
         values (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s)"""
usermaster.executemany(insertsql, spamlog)
platdb.commit()
platdb.close()
spamlog.close()
conn.close()
