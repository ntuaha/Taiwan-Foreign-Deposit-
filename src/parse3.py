# -*- coding: utf-8 -*- 

import xlrd
import re

#處理掉unicode 和 str 在ascii上的問題
import sys 
import os
reload(sys) 
sys.setdefaultencoding('utf8') 

def parse(source_path,destination_path):
	data_month = {}
	data_day = {}	
	for num in range(1,8):
		book = xlrd.open_workbook(source_path+num+".xls")
		#print "The number of worksheets is",book.nsheets
		#print "Worksheet name(s):", book.sheet_names()
		#header = ['年月','銀行','銀行英文','全行外匯活期存款','全行外匯定期存款','全行總額','國內外匯活期存款','國內外匯定期存款','國內總額','海外外匯活期存款','海外外匯定期存款','海外總額','金控註記']
		total_data = range(4)
		print sh.name, sh.nrows, sh.ncols
		for i in range(2,sh.nrows):
			row_name = unicode(sh.cell_value(rowx=i,colx = 0))
			n = row_name.split('/')
			year = n[0]
			month = n[1]
			day = n[2]
			data_day[year] ={month:{day:[float(sh.cell_value(rowx=i,colx = 2)),float(sh.cell_value(rowx=i,colx = 3))]}}

	for y in data_day:
		data_month[y] = {}
		for m in data_day[y]:
			current_day = 32
			for d in data_day[y][m]:
				if data_month[y][m] == None:
					data_month[y][m] = data_day[y][m][d]
					current_day = d
				else:
					if current_day > d:
						data_month[y][m] = data_day[y][m][d]




	


	output(destination_path,data_day,data_month)
				







#輸出
def output(destination_path,data_day,data_month):
	f = open("%sExangeDay.csv"%(destination_path),"a+")
	for y in data_day.keys().sort():
		for m in data_day[y].keys().sort():
			for d in data_day[y][m].keys().sort():
				d = [y,m,d,data[y][m][d][0],data[y][m][d][1]]
				f.write(",".join(d)+"\n")
	f.close()
	f = open("%sExangeMonth.csv"%(destination_path),"a+")
	for y in data_month.keys().sort():
		for m in data_month[y].keys().sort():
				d = [y,m,d,data[y][m][0],data[y][m][d][1]]
				f.write(",".join(d)+"\n")
	f.close()


 		




if __name__ == '__main__':
	from_path= '/Users/aha/Dropbox/Project/Financial/Data/exchange/'
	to_path = '/Users/aha/Dropbox/Project/Financial/Codes/'
	cmd = "rm %sExangeDay.csv"%(to_path)
	os.system(cmd)
	cmd = "rm %sExangeMonth.csv"%(to_path)
	os.system(cmd)


	f = open("%sExangeDay.csv"%(to_path),"a+")
	total_header = ['年','月','日','買進','賣出']
	f.write(",".join(total_header)+"\n")
	f.close()
	f = open("%sExangeMonth.csv"%(to_path),"a+")
	total_header = ['年','月','買進','賣出']
	f.write(",".join(total_header)+"\n")
	f.close()

	file_num  = 1
	#for yy in range(95,100):
#		for mm in range(1,13):

	parse(from_path,to_path,file_num)
	#parse(from_path,to_path,'10001')


