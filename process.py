import os, re, sys, time, xlrd, pyodbc, datetime
from datetime import date
import fnmatch
import numpy as np
import pandas as pd
import itertools as it
from openpyxl import load_workbook
from shutil import copyfile
import myfun as dd
import dbquery as dbq

#------ program starting point --------	
if __name__=="__main__":
	## dd/mm/yyyy format
	print 'Process date is ' + str(time.strftime("%m/%d/%Y"))
	print 'Please enter the cycle end date (mm/dd/yyyy) you want to process:'
	#-----------------------------------------------------
	#------- get cycle date ----------------------
	getcycledate = datetime.datetime.strptime(raw_input(), '%m/%d/%Y')
	endday = getcycledate
	startday = datetime.datetime.strptime('1/1/' + str(endday.year), '%m/%d/%Y')
	supportalyear = endday.year + 1

	print 'Cycle start date is ' + str(startday)
	print 'Cycle end date is ' + str(endday)
	#--------- database info ----------------
	driver = r"{Microsoft Access Driver (*.mdb, *.accdb)};"

	db_file = r"F:\\Files For\\Hai Yen Nguyen\\Practice Credits\\PC.accdb;"
	#db_file = r"C:\\pycode\\CsltPC\\PC.accdb;"
	user = "admin"
	password = ""
	#--------------------------------------------------------------------
	#--------- get list of cslts with support AL and New Business/Managed Asset ----------
	sql = '''
			SELECT 
				qry_CsltList.[Cslt]
				,qry_CsltList.[Name]
				,qry_CsltList.[Status]
				,qry_CsltList.[Position]
				,qry_CsltList.[TermDate]
				,qry_CsltList.[SupportAL]
				,qry_CsltList.[BusCrdt]
				,qry_CsltList.[MgmtAsst]
			FROM qry_CsltList
			WHERE (qry_CsltList.[EYear]) = ''' + str(supportalyear) + '''
			ORDER BY 	
				qry_CsltList.[Cslt]
		'''

	dfpc = dbq.df_select(driver, db_file, sql)
	dfpc['SupportAL'] = dfpc['SupportAL'].values.astype(np.int64)
	#----------- get current year rate for practice credit ------
	sql = '''
			SELECT PracticeCreditRate.AL AS [SupportAL]
				,PracticeCreditRate.LumpSum
				,PracticeCreditRate.BusCrdtRate
				,PracticeCreditRate.BusCrdtQty
				,PracticeCreditRate.MgmtAsstRate
				,PracticeCreditRate.MgmtAsstQty
			FROM PracticeCreditRate
			WHERE ([PracticeCreditRate].[Period]) = #''' + endday.strftime("%m/%d/%Y") + '''#
		'''

	dfpcrate = dbq.df_select(driver, db_file, sql)
#	dfpcrate['SupportAL'] = dfpcrate['SupportAL'].values.astype(np.int64)
#	dfpcrate['LumpSum'] = dfpcrate['LumpSum'].values.astype(np.float64)
#	dfpcrate['BusCrdtRate'] = dfpcrate['BusCrdtRate'].values.astype(np.float64)
#	dfpcrate['BusCrdtQty'] = dfpcrate['BusCrdtQty'].values.astype(np.float64)
#	dfpcrate['MgmtAsstRate'] = dfpcrate['MgmtAsstRate'].values.astype(np.float64)
#	dfpcrate['MgmtAsstQty'] = dfpcrate['MgmtAsstQty'].values.astype(np.float64)

	dfpc = dfpc.merge(dfpcrate, on='SupportAL', how='left')
	
	#dfpc['BusCrdtAmt'] = dfpc['BusCrdt'] / dfpc['BusCrdtQty'] * dfpc['BusCrdtRate']
	#dfpc['MgmtAsstAmt'] = dfpc['MgmtAsst'] / dfpc['MgmtAsstQty'] * dfpc['MgmtAsstRate']
	#dfpc['PCAmt'] = dfpc['LumpSum'] + dfpc['BusCrdtAmt'] + dfpc['MgmtAsstAmt']

	dfpc.loc[dfpc['Status'] == 'Active', 'BusCrdtAmt'] = dfpc['BusCrdt'] / dfpc['BusCrdtQty'] * dfpc['BusCrdtRate']
	dfpc.loc[dfpc['Status'] == 'Active', 'MgmtAsstAmt'] = dfpc['MgmtAsst'] / dfpc['MgmtAsstQty'] * dfpc['MgmtAsstRate']
	dfpc.loc[dfpc['Status'] != 'Active', ['LumpSum', 'BusCrdtRate', 'BusCrdtQty', 'MgmtAsstRate', 'MgmtAsstQty']] = np.NAN
	dfpc['PCAmt'] = dfpc['LumpSum'] + dfpc['BusCrdtAmt'] + dfpc['MgmtAsstAmt']
	
	
	#--------- output to Excel ---------------------
	writer = pd.ExcelWriter('PC.xlsx', engine='xlsxwriter')
	
	dfpc.to_excel(writer, sheet_name='PC', startrow=1, freeze_panes=(2,8), index=False)

	workbook = writer.book
	worksheet = writer.sheets['PC']
	
	# Add some cell formats.
	formatcell = workbook.add_format({'bold': True, 'align':'center'})
	formatcslt = workbook.add_format({'bold':True, 'bg_color':'#FFFF00'})
	formatdate = workbook.add_format({'num_format':'mm/dd/yyyy'})
	formatnum = workbook.add_format({'num_format':'#,##0.00'})
	formatpercent = workbook.add_format({'num_format':'0.00%'})
	formatbi = workbook.add_format({'num_format':'#,##0.00', 'bold':True, 'bg_color':'#FFFF00'})

#	#----------- add notes --------------
#	worksheet.write('P1', 'N-O', formatcell)
#	worksheet.write('S1', 'Q+R', formatcell)
#	worksheet.write('V1', 'N*U', formatcell)
#	worksheet.write('W1', 'V/5', formatcell)
#	worksheet.write('Z1', 'X*R', formatcell)
#	worksheet.write('AA1', 'N*Y', formatcell)
#	worksheet.write('AB1', 'Z+AA', formatcell)
#	
	# Set the column width and format
	worksheet.set_column('A:A', 8, formatcslt)
	worksheet.set_column('B:B', 15)
	worksheet.set_column('E:E', 10, formatdate)
	worksheet.set_column('F:F', 10, formatcslt)
	worksheet.set_column('G:I', 14, formatnum)
	worksheet.set_column('J:J', 14, formatpercent)
	worksheet.set_column('K:O', 14, formatnum)
	worksheet.set_column('P:P', 14, formatbi)

	# Close the Pandas Excel writer and output the Excel file.
	writer.save()
	
	print 'The process is done'
	
	
	