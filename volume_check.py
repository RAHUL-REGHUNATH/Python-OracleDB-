"""
Author : Rahul Reghunath (EBA - ETL)
Date   : 30 - 04 -2020
"""

import cx_Oracle
import win32com.client as win32
import pandas as pd
import datetime

now = datetime.datetime.now()


def DB_Steps(connect):
    connection = connect.cursor()
    try:
        print("Excecuting Query....")
		# Enter your query inside 
        connection.execute("""
            select 'Sales Order',count(*) from work.wrk_so_line_dly
            union all
            select 'AR invoice',count(*) from work.wrk_ar_invoice_line
            union all
            select 'AR',count(*) from STG_ERP.RA_CUSTOMER_TRX_LINES_ALL where ETL_REPROCESS_FLAG='N'
            union all
            select 'GL',count(*) from STG_ERP.RA_CUST_TRX_LINE_GL_DIST_ALL where ETL_REPROCESS_FLAG='N'
            union all
            select 'IB',count(*) from STG_SAP.ZRIBASE where etl_reprocess_flag ='N'
            union all
            select 'GS',count(*) "GS_CUSTOMER_ORDER_HEADER" from STG_SAP.GS_CRMD_ORDERADM_H
            union all
            select 'QUOTE_HEADERS_ALL',count(*) from STG_ERP_R12.ASO_QUOTE_HEADERS_ALL where ETL_REPROCESS_FLAG='N'
            union all
            select 'QUOTE_LINES_ALL',count(*) from STG_ERP_R12.ASO_QUOTE_LINES_ALL where ETL_REPROCESS_FLAG='N'
            union all
            select 'Sales Order INCR',count(*) from work.wrk_so_line_incr
            union all
            SELECT 'RevPro1',COUNT(*) FROM STG_ERP.RPRO_RC_SCHD_G WHERE ETL_REPROCESS_FLAG='N'
            union all
            SELECT 'RevPro2',COUNT(*) FROM STG_ERP.RPRO_RC_LINE_G WHERE ETL_REPROCESS_FLAG='N'
            union all
            SELECT 'RevPro3',COUNT(*) FROM STG_ERP.RPRO_RC_BILL_G WHERE ETL_REPROCESS_FLAG='N'
            union all
            SELECT 'RevPro4',COUNT(*) FROM STG_ERP.RPRO_RC_SCHD_DEL_G WHERE ETL_REPROCESS_FLAG='N'
             """)

        result = connection.fetchall()
        Mail(result)

    except cx_Oracle.DatabaseError as errors:
        print(errors)

    finally:
        connection.close()
        connect.close()
        print('DB closed')


def Mail(result):
    print("Preparing mail....")
	
	# baseline volume count
    table = pd.DataFrame(result, columns=['Source', 'Count'])
    average = [15000, 5000, 15000, 100000, 25000, 45000, 30000, 300000, 20000,200000,50000,50000,250000]
    table['Baseline'] = average
    Status = []
    for index in range(13):
        if int(table['Count'][index]) > (average[index]):
            Status.insert(index, 'High')
        else:
            Status.insert(index, 'Low')
    table['Status'] = Status

	# sending mail
    outlook = win32.Dispatch('outlook.application')
    mail = outlook.CreateItem(0)
    mail.To = input("To which DL mail to be sent:")
    mail.Subject = 'VOLUME DETAILS : ' + now.strftime("%Y-%m-%d %H:%M:%S")
    print('Adding contents to mail....')
    mail.HTMLBody = '<html><body>' + table.to_html() + '</body></html>'
    mail.Send()
    print("--Mail sent successfully--")


if __name__ == '__main__':

    try:
        print('--DB steps--')
        # user = input("Enter your DB Username : ")
        # password = input("Enter your DB password : ")
        dsn_tns = cx_Oracle.makedsn('enter host name', 'enter port number', service_name='enter service name')
        conn = cx_Oracle.connect(user='enter username', password='enter password', dsn=dsn_tns) 
        print("--Connection Successful--")
        DB_Steps(conn)
    except cx_Oracle.DatabaseError as e:
        print(e)


