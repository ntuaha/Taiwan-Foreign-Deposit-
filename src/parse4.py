# -*- coding: utf-8 -*- 

import re

#處理掉unicode 和 str 在ascii上的問題
import sys 
import os
reload(sys) 
sys.setdefaultencoding('utf8') 


def read(f,bank):
	for row in f:
		cols = row.strip().split(',')

		if unicode(cols[1]).find(bank) > -1:
			return cols

def parse(source_path,destination_path,bank):
	data = {}
	for year in range(95,102):
		if year not in data:
			data[year]={}
		for month in range(1,13):
			row = read(open(source_path+"%d%02d"%(year,month)+".csv","r"),bank[1])
			data[year][month] = row
	#print data
	data[102] = {}
	row = read(open(source_path+"%d%02d"%(102,1)+".csv","r"),bank[1])
	data[102][1] = row
	row = read(open(source_path+"%d%02d"%(102,2)+".csv","r"),bank[1])
	data[102][2] = row

	output(destination_path,data,bank)
				







#輸出
def output(destination_path,data,header):
	f = open("%s%s.csv"%(to_path,bank[0]),"w+")
	for year in sorted(data):
		for month in sorted(data[year]):
			#print data[year][month]
			if data[year][month] != None:
				d = [year,month] + data[year][month][3:12]
				f.write(",".join(map(str,d))+"\n")
	f.close()


 		




if __name__ == '__main__':
	from_path= '/Users/aha/Dropbox/Project/Financial/Codes/csv/'
	to_path = '/Users/aha/Dropbox/Project/Financial/Codes/csv/'
	#bank=['ESun',u'玉山']
	#bank=['Taishin',u'台新']
	#bank=['SinoPac',u'永豐']
	#bank=['Citibank',u'花旗']
	bank=['Taiwan',u'臺灣銀行']
	cmd = "rm %s%s.csv"%(to_path,bank[0])
	os.system(cmd)

	f = open("%s%s.csv"%(to_path,bank[0]),"w+")
	total_header = ['年月','全行外匯活期存款','全行外匯定期存款','全行總額','國內外匯活期存款','國內外匯定期存款','國內總額','海外外匯活期存款','海外外匯定期存款','海外總額']
	f.write(",".join(total_header)+"\n")
	f.close()

	parse(from_path,to_path,bank)


