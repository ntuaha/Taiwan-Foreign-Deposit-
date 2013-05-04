# -*- coding: utf-8 -*- 

import xlrd
import re
import sys 
reload(sys) 
sys.setdefaultencoding('utf8') 

def parse(source_path,destination_path,date):
	f = open(destination_path+date+"_raw.csv","w+")
	book = xlrd.open_workbook(source_path+date+".xls")
	print "The number of worksheets is",book.nsheets
	print "Worksheet name(s):", book.sheet_names()
	sh = book.sheet_by_index(0)
	print sh.name, sh.nrows, sh.ncols
	#for i in range(sh.nrows):
#		print "Line %d" %i	
#		print " ".join([unicode(sh.cell_value(rowx=i,colx=j)) for  j in range(sh.ncols)])
		#for j in range(sh.ncols):
		#	print sh.cell_value(rowx=i,colx=j)
	#fopen
	header = ['年月','銀行','銀行英文','全行外匯活期存款','全行外匯定期存款','全行總額','國內外匯活期存款','國內外匯定期存款','國內總額','海外外匯活期存款','海外外匯定期存款','海外總額','金控註記']
#	print header

	bank_data = {}
	total_data = range(len(header)-3)
	jump_gap = 8
	mode  = 0
	for i in range(sh.nrows):
		row_name = unicode(sh.cell_value(rowx=i,colx = 0))

		#print u"2-5　一般銀行外匯存款餘額									"
		#if row_name == u"2-5　一般銀行外匯存款餘額":
		#	print "yes"
		if unicode(sh.cell_value(rowx=i,colx = 1)) == u"":	
			#空的但是資料開頭就跳到資料頭		
			if  row_name == u"2-5　一般銀行外匯存款餘額":
				mode =1
			
			#if row_name == u"2-5　一般銀行外匯存款餘額（續一）":
	#			mode1_line = i+jump_gap
			
	#		if row_name== u"2-5　一般銀行外匯存款餘額（續二）":
	#			mode1_line = i+jump_gap
			
			if row_name == u"2-5　一般銀行外匯存款餘額（續三）":
				mode =2

			print "%d: Empty:%s" %(i,row_name)
			continue
		
		#第二藍衛如果是文字就跳過
		if len(re.findall(r"[-+]?\d*\.\d+|\d+",unicode(sh.cell_value(rowx=i,colx = 1)))) == 0:
			continue

		#全行總和
		if row_name== u"總　　　　　計　Total" and 1 == mode:
			total_data[0] = date
			total_data[1] = float(sh.cell_value(rowx=i,colx = 1))*1e6
			total_data[2] = float(sh.cell_value(rowx=i,colx = 2))*1e6
			total_data[3] = float(sh.cell_value(rowx=i,colx = 3))*1e6
			continue
		#全行銀行
		if 1 == mode:
			bank_name = unicode(sh.cell_value(rowx=i,colx = 0))
			bank_data[bank_name] = {}
			bank_data[bank_name]["ALL_MY"] = float(sh.cell_value(rowx=i,colx = 1))*1e6
			bank_data[bank_name]["ALL_FY"] = float(sh.cell_value(rowx=i,colx = 2))*1e6
			bank_data[bank_name]["ALL_Y"] = float(sh.cell_value(rowx=i,colx = 3))*1e6


		#全行總和
		if sh.cell_value(rowx=i,colx = 0)== u"總　　　　　計　Total" and 2 == mode:
			#國內
			total_data[4] = float(sh.cell_value(rowx=i,colx = 1))*1e6
			total_data[5] = float(sh.cell_value(rowx=i,colx = 2))*1e6
			total_data[6] = float(sh.cell_value(rowx=i,colx = 3))*1e6
			#Oversea 海外
			total_data[7] = total_data[1] - total_data[4]
			total_data[8] = total_data[2] - total_data[5]
			total_data[9] = total_data[3] - total_data[6]
			continue
		#全行銀行
		if 2 == mode:
			bank_name = unicode(sh.cell_value(rowx=i,colx = 0))
			bank_data[bank_name]["DB_MY"] = float(sh.cell_value(rowx=i,colx = 1))*1e6
			bank_data[bank_name]["DB_FY"] = float(sh.cell_value(rowx=i,colx = 2))*1e6
			bank_data[bank_name]["DB_Y"] = float(sh.cell_value(rowx=i,colx = 3))*1e6
			bank_data[bank_name]["OS_MY"] = bank_data[bank_name]["ALL_MY"] - bank_data[bank_name]["DB_MY"]
			bank_data[bank_name]["OS_FY"] = bank_data[bank_name]["ALL_FY"] - bank_data[bank_name]["DB_FY"]
			bank_data[bank_name]["OS_Y"] = bank_data[bank_name]["ALL_Y"] - bank_data[bank_name]["DB_Y"]
				#print "%s %% %s" %(bank_data[bank_name],bank_name)

				

#輸出
	f = open("%sTotal_%s.csv"%(destination_path,date),"w+")
	
	total_header = ['年月','全行外匯活期存款','全行外匯定期存款','全行總額','國內外匯活期存款','國內外匯定期存款','國內總額','海外外匯活期存款','海外外匯定期存款','海外總額']
	f.write(",".join(total_header)+"\n")
	f.write(",".join(map(str,map(int,total_data))))
	f.close()

	f = open("%s%s.csv"%(destination_path,date),"w+")
	header = ['年月','銀行','銀行英文','全行外匯活期存款','全行外匯定期存款','全行總額','國內外匯活期存款','國內外匯定期存款','國內總額','海外外匯活期存款','海外外匯定期存款','海外總額','金控註記']
	f.write(",".join(header)+"\n")
	for i in bank_data:
		try:
			d = [None]*13
			d[0] = date
			d[1] = i
			d[2] = ""
			d[3] = int(bank_data[i]["ALL_MY"])
			d[4] = int(bank_data[i]["ALL_FY"])
			d[5] = int(bank_data[i]["ALL_Y"])
			d[6] = int(bank_data[i]["DB_MY"])
			d[7] = int(bank_data[i]["DB_FY"])
			d[8] = int(bank_data[i]["DB_Y"])
			d[9] = int(bank_data[i]["OS_MY"])
			d[10] = int(bank_data[i]["OS_FY"])
			d[11] = int(bank_data[i]["OS_Y"])
			d[12] = 0
			f.write(",".join(map(str,d))+"\n")
		except KeyError:
			f.write(",".join(map(str,d))+"\n")

	f.close()


 		




if __name__ == '__main__':
	parse('/Users/aha/Dropbox/Project/Financial/Data/','','9501')
