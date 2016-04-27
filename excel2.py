from win32com.client import Dispatch 
from xml.dom import minidom
import os
import win32com.client as win32
import time
import sets
class excel2(object):
	
	device=[]
	
	def __init__(self,dc1,xlwb1,Hi,Me,Lo):
	
		
		
		self.High=Hi
		self.Medium=Me
		self.Low=Lo
		word = Dispatch('Word.Application')
		xlApp = Dispatch("Excel.Application")
		xlApp.Visible = True
		
	
		self.doc1 = dc1
		
		
		self.xlWb1 =xlwb1
		self.xlSht1 = self.xlWb1.WorkSheets(3)
		
		
		print "Please give path of second excel sheet"
		print "\n"
		self.xlWb2 =xlApp.Workbooks.Open(raw_input())
		
		self.xlSht2 = self.xlWb2.WorkSheets(3)
		print "Please give path and name to save file e.g C:\Users\inprsha\Desktop\hello.xlsx"
		print "\n"
		self.xlWb2.SaveAs(raw_input())
		print "enter row no in which u want to fill details"
		self.row_no=raw_input()
		# xlApp.Application.Quit()
	

	def fetchdvcdetails(self):
		rngDoc = self.doc1.Range(0, 0)
		table = self.doc1.Tables(1)
		
		self.device.append(table.Cell(Row=1, Column=2).Range.Text.split('.')[1])#device name  
		self.device.append(table.Cell(Row=1, Column=2).Range.Text[:-1].split('.')[2][:-1])#test run
		self.device.append(table.Cell(Row=3, Column=2).Range.Text[:-1])#test end date
		self.device.append(table.Cell(Row=4, Column=2).Range.Text[:-1])#last modification of report
		self.device.append(table.Cell(Row=5, Column=2).Range.Text[:-1])#testers name
		table = self.doc1.Tables(4)
		self.device.append(table.Cell(Row=5, Column=2).Range.Text[:-1])#os
		self.device.append(table.Cell(Row=4, Column=2).Range.Text[:-1])#firmware version
		table = self.doc1.Tables(2)
		self.device.append(table.Cell(Row=2, Column=2).Range.Text[:-1])#owner
		self.device=filter(lambda name: name.strip(), self.device)
		
		
	def filldvcdetails(self):
	
		if self.device[0].find("Regression")== -1 :
			self.xlSht2.Cells(self.row_no,3).Value=self.device[0]
		else:
			self.xlSht2.Cells(self.row_no,3).Value=self.device[0].split('Regression')
		# self.xlSht2.Cells().Value=self.device[1])
		self.xlSht2.Cells(self.row_no,10).Value=self.device[2]
		
		self.xlSht2.Cells(self.row_no,20).Value=self.device[3]
		self.xlSht2.Cells(self.row_no,36).Value=self.device[4]
		self.xlSht2.Cells(self.row_no,17).Value=self.device[5]
		self.xlSht2.Cells(self.row_no,4).Value=self.device[6]
		self.xlSht2.Cells(self.row_no,8).Value=self.device[7]
		self.xlSht2.Cells(self.row_no,5).Value=time.strftime("%d/%m/%Y")
		self.xlSht2.Cells(self.row_no,14).Value=self.device[6]
	def NoOfFauts(self):
		# print self.High
		# print self.Medium
		self.xlSht2.Cells(self.row_no,22).Value=self.High
		self.xlSht2.Cells(self.row_no,24).Value=self.Medium
		self.xlSht2.Cells(2,30).Value="No. of LOW severity level FAULTS"
		self.xlSht2.Cells(self.row_no,30).Value=self.Low
		
	def testtype(self):
		
		if self.device[0].find("Regression")== -1 :
			# print self.device[1]
			if self.device[1]== "Test1" :
				self.xlSht2.Cells(self.row_no,6).Value="Initial"
			else:
				self.xlSht2.Cells(self.row_no,6).Value="Retest"
		else:
			self.xlSht2.Cells(self.row_no,6).Value="Regression"
		
	def devarrival(self):
		
		xmlfile = minidom.parse("C:/Users/inprsha/Desktop/nmap.xml")
		nrun =xmlfile.getElementsByTagName('nmaprun')
		for attbr in nrun:
			startstr=attbr.getAttribute('startstr')
		# print str(startstr)
		self.xlSht2.Cells(self.row_no,11).Value=str(startstr).split(' ')[1] +" " +str(startstr).split(' ')[3]+" " +str(startstr).split(' ')[5]
		self.xlSht2.Cells(self.row_no,13).Value=str(startstr).split(' ')[1] +" " +str(startstr).split(' ')[3]+" " +str(startstr).split(' ')[5]
	
	def testedprotocol(self):
	
		testedprotocols=[]
		
		if self.xlSht2.Cells(self.row_no,6).Value=="Regression":
			print "fdlkjdlfjlf"
			for i in range (9,14):
				table=self.doc1.Tables(i)
				for rows in range(2,self.doc1.Tables(i).Rows.Count+1):
					try:
						testedprotocols.append(table.Cell(Row=rows, Column=1).Range.Text[:-1])# appending  tested protocols in table 5,6,7,8,9 of word template
					except:
						print ""
			testedprotocols=filter(lambda name: name.strip(),testedprotocols)
			tstprotocol = sets.Set(testedprotocols)#to remove duplicates in array
			# print tstprotocol
			self.xlSht2.Cells(self.row_no,15).Value='/'.join(tstprotocol)
			
			
		else:
			testedprotocols.append("ARP/IP/ICMP/TCP/UDP")
			self.xlSht2.Cells(self.row_no,15).Value=testedprotocols
			
		
if __name__ == "__main__":
	r =excel2()
	r.fetchdvcdetails()
	r.filldvcdetails()
	r.NoOfFauts()
	r.testtype()
	r.devarrival()
	r.testedprotocol()
	# print "=============================="
	# print r.faults[2]
	# # print r.High
	# print "=============================="
	# print r.faults[3]
	# # print r.Medium
	# # print "=============================="
	# # print r.Low
	# # print "=============================="
	
	

