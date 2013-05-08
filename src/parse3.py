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
		print source_path+str(num)+".xls"
		book = xlrd.open_workbook(source_path+str(num)+".xlsx")
		sh = book.sheet_by_index(0)
		#print "The number of worksheets is",book.nsheets
		#print "Worksheet name(s):", book.sheet_names()
		#header = ['年月','銀行','銀行英文','全行外匯活期存款','全行外匯定期存款','全行總額','國內外匯活期存款','國內外匯定期存款','國內總額','海外外匯活期存款','海外外匯定期存款','海外總額','金控註記']
		total_data = range(4)
		print sh.name, sh.nrows, sh.ncols
		for i in range(2,sh.nrows):
			row_name = unicode(sh.cell_value(rowx=i,colx = 0))
			n = row_name.strip().split('/')
			print n
			year = int(n[0])
			month = int(n[1])
			day = int(n[2])
			if year not in data_day:
				data_day[year] = {}
			if month not in data_day[year]:
				data_day[year][month] = {}
			if day not in data_day[year][month]:
				data_day[year][month][day] = [float(sh.cell_value(rowx=i,colx = 2)),float(sh.cell_value(rowx=i,colx = 3))]

	for y in data_day:
		data_month[y] = {}
		for m in data_day[y]:
			current_day = 32
			if m not in data_month[y]:
				data_month[y][m] = []
			for d in data_day[y][m]:
				if data_month[y][m] == []:
					data_month[y][m] = [d,data_day[y][m][d][0],data_day[y][m][d][1]]
					current_day = d
				else:
					if current_day > d:
						data_month[y][m] = [d,data_day[y][m][d][0],data_day[y][m][d][1]]




	

	#print data_day
	#print data_month
	output(destination_path,data_day,data_month)
				







#輸出
def output(destination_path,data_day,data_month):
	f = open("%sExangeDay.csv"%(destination_path),"a+")
	for y in sorted(data_day.keys()):
		for m in sorted(data_day[y].keys()):
			for d in sorted(data_day[y][m]):
				dd = map(str,[y,m,d,data_day[y][m][d][0],data_day[y][m][d][1]])
				print dd
				f.write(",".join(dd)+"\n")
	f.close()
	f = open("%sExangeMonth.csv"%(destination_path),"a+")
	for y in sorted(data_day):
		for m in sorted(data_day[y]):			
			d = map(str,[y,m,data_month[y][m][0],data_month[y][m][1],data_month[y][m][2]])
			f.write(",".join(d)+"\n")
	f.close()


 		




if __name__ == '__main__':
	from_path= '/Users/aha/Dropbox/Project/Financial/Data/exchange/'
	to_path = '/Users/aha/Dropbox/Project/Financial/Codes/'
	cmd = "rm %sExangeDay.csv"%(to_path)
	os.system(cmd)
	cmd = "rm %sExangeMonth.csv"%(to_path)
	os.system(cmd)


	f = open("%sExangeDay.csv"%(to_path),"w+")
	total_header = ['年','月','日','買進','賣出']
	f.write(",".join(total_header)+"\n")
	f.close()
	f = open("%sExangeMonth.csv"%(to_path),"w+")
	total_header = ['年','月','日','買進','賣出']
	f.write(",".join(total_header)+"\n")
	f.close()

	#for yy in range(95,100):
#		for mm in range(1,13):

	parse(from_path,to_path)
	#parse(from_path,to_path,'10001')


