# -*- coding: utf-8 -*- 

import xlrd

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
	read_flag = 0 #不作用
	jump_gap = 8
	mode1_line = -1
	mode2_line = -1

	for i in range(sh.nrows):
		row_name = unicode(sh.cell_value(rowx=i,colx = 0))
		#print row_name
		#print u"2-5　一般銀行外匯存款餘額									"
		#if row_name == u"2-5　一般銀行外匯存款餘額":
		#	print "yes"
		if unicode(sh.cell_value(rowx=i,colx = 1)) == u"":	
			#空的但是資料開頭就跳到資料頭		
			if  row_name == u"2-5　一般銀行外匯存款餘額":
				read_flag = 1 # 全行
				mode1_line = i+jump_gap
				print "HI %d" %mode1_line
			
			if row_name == u"2-5　一般銀行外匯存款餘額（續一）":
				read_flag = 1 # 全行
				mode1_line = i+jump_gap
			
			if row_name== u"2-5　一般銀行外匯存款餘額（續二）":
				read_flag = 1 # 全行
				mode1_line = i+jump_gap
			
			if row_name == u"2-5　一般銀行外匯存款餘額（續三）":
				read_flag = 2 # 國內
				mode2_line = i+jump_gap
			
			if row_name == u"2-5　一般銀行外匯存款餘額（續四）":
				read_flag = 2 # 國內
				mode2_line = i+jump_gap
			
			if row_name == u"2-5　一般銀行外匯存款餘額（續五完）":
				read_flag = 2 # 國內
				mode2_line = i+jump_gap
		else:
			#不是空的就讀下一行
			if i>read1_line
			mode1_line =  mode1_line+1
			mode2_line =  mode2_line+1
		print "===mode1_line %d"%mode1_line
		print "mode2_line %d"%mode2_line	
		print "%d  || %s" % (i,unicode(sh.cell_value(rowx=i,colx = 1)))
		#print bank_data
		#print row_name



		#全行總和
		if row_name== u"總　　　　　計　Total" and i == mode1_line:
			total_data[0] = date
			total_data[1] = float(sh.cell_value(rowx=i,colx = 1))*1e6
			total_data[2] = float(sh.cell_value(rowx=i,colx = 2))*1e6
			total_data[3] = float(sh.cell_value(rowx=i,colx = 3))*1e6
		#全行銀行
		if i == mode1_line:
			bank_name = unicode(sh.cell_value(rowx=i,colx = 0))
			bank_data[bank_name] = {}
			
			
			try:
				bank_data[bank_name]["ALL_MY"] = float(sh.cell_value(rowx=i,colx = 1))*1e6
				bank_data[bank_name]["ALL_FY"] = float(sh.cell_value(rowx=i,colx = 2))*1e6
				bank_data[bank_name]["ALL_Y"] = float(sh.cell_value(rowx=i,colx = 3))*1e6
				print "%s %% %s" %(bank_data[bank_name],bank_name)
			except ValueError:
				print "NO"
				bank_data[bank_name]["ALL_MY"] = -1
				bank_data[bank_name]["ALL_FY"] = -1
				bank_data[bank_name]["ALL_Y"] = -1
			print unicode(sh.cell_value(rowx=i,colx = 0))
			print "%s %% %s" %(bank_data[bank_name],bank_name)
		#全行總和
		if sh.cell_value(rowx=i,colx = 0)== u"總　　　　　計　Total" and i == mode2_line:
			#國內
			total_data[4] = float(sh.cell_value(rowx=i,colx = 1))*1e6
			total_data[5] = float(sh.cell_value(rowx=i,colx = 2))*1e6
			total_data[6] = float(sh.cell_value(rowx=i,colx = 3))*1e6
			#Oversea 海外
			total_data[7] = total_data[1] - total_data[4]
			total_data[8] = total_data[2] - total_data[5]
			total_data[9] = total_data[3] - total_data[6]
		#全行銀行
		if i == mode2_line:
			bank_name = unicode(sh.cell_value(rowx=i,colx = 0))
			try:	
				bank_data[bank_name]["DB_MY"] = float(sh.cell_value(rowx=i,colx = 1))*1e6
				bank_data[bank_name]["DB_FY"] = float(sh.cell_value(rowx=i,colx = 2))*1e6
				bank_data[bank_name]["DB_Y"] = float(sh.cell_value(rowx=i,colx = 3))*1e6
				bank_data[bank_name]["OS_MY"] = bank_data[bank_name]["ALL_MY"] - bank_data[bank_name]["OS_MY"]
				bank_data[bank_name]["OS_FY"] = bank_data[bank_name]["ALL_FY"] - bank_data[bank_name]["OS_FY"]
				bank_data[bank_name]["OS_Y"] = bank_data[bank_name]["ALL_Y"] - bank_data[bank_name]["OS_Y"]
				print "%s %% %s" %(bank_data[bank_name],bank_name)
			except ValueError:
				print "NO"
				bank_data[bank_name]["DB_MY"] = -1
				bank_data[bank_name]["DB_FY"] = -1
				bank_data[bank_name]["DB_Y"] = -1
				bank_data[bank_name]["OS_MY"] = -1
				bank_data[bank_name]["OS_FY"] = -1
				bank_data[bank_name]["OS_Y"] = -1
			print bank_name
			print unicode(sh.cell_value(rowx=i,colx = 0))
			print "%s %% %s" % (bank_data[bank_name],bank_name)
	print "bank_data ===="
	print bank_data
	print "total_data ===="
	print total_data





 		




if __name__ == '__main__':
	parse('/Users/aha/Dropbox/Project/Financial/Data/','','9501')
