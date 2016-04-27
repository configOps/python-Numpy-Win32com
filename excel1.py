# from win32com.client import Dispatch
# xlApp = Dispatch("Excel.Application")
# xlWb1 = xlApp.Workbooks.Open("C:\Users\inprsha\Desktop\PA_Test_Detail_Faults.xlsx")

# xlSht1 = xlWb1.WorkSheets(3)
# xlSht1.Visible = True
# print xlSht1.Cells(3,2).Value

# xlApp.Application.Quit()

from win32com.client import Dispatch 
import os
import win32com.client as win32
from excel2 import excel2

class excel1(object):

# This class fetches Device Details from MS Word Template and write in MS Excel Template

	faults=[]
	High = []
	Medium = []
	Low = []
	device=[]
	excel2Object=[]
	
	def __init__(self):
		word = Dispatch('Word.Application')
		xlApp = Dispatch("Excel.Application")
		xlApp.Visible = True
		
		print "Please give path of Ms WOrd report template "
		self.doc1 = word.Documents.Open(raw_input())
		
		print "Please give path and name of excel file having  fAULTS details e.g C:\Users\inprsha\Downloads\PA_Test_Detail_Faults.xlsx"
		print "\n"
		self.xlWb1 =xlApp.Workbooks.Open(raw_input())
		
		self.xlSht1 = self.xlWb1.WorkSheets(3)
		
		print "Please give path and name to save file e.g C:\Users\inprsha\Desktop\hello.xlsx"
		print "\n"
		self.xlWb1.SaveAs(raw_input())
		# xlApp.Application.Quit()
	
	def fetchpnum(self,searchable):
		# print self.doc1.Paragraphs.Count
		for i in range(1, self.doc1.Paragraphs.Count):
			value = self.doc1.Paragraphs(i).Range.Text.lower()
			if value.startswith(searchable.lower()):
				return i
		return 0
		
	def highfaults(self):
		start=self.fetchpnum('High severity faults in detail')
		end=self.fetchpnum('Medium severity faults in detail')
		for i in range(start+1,end):
			value=self.doc1.Paragraphs(i).Range.Text
			if value[:-1].lower()!="No High severity faults were observed.".lower():
				
				self.High.append(value)
		self.High=filter(lambda name: name.strip(), self.High)		
		self.faults.append(self.High)
		# print "hello from fnctn"
		# print self.High
		# print "=============================="
	
			
	def mediumfaults(self):
		start=self.fetchpnum('Medium severity faults in detail')+1
		end=self.fetchpnum('Low severity faults in detail')
		
		for i in range(start,end):
			value=self.doc1.Paragraphs(i).Range.Text
			if value[:-1].lower()!="No Medium severity faults were observed.".lower():
				
				self.Medium.append(value)
		self.Medium=filter(lambda name: name.strip(), self.Medium)
		self.faults.append(self.Medium)
	
	
		
		
	def lowfaults(self):
		start=self.fetchpnum('Low severity faults in detail')
		end=self.fetchpnum('Comparison with respect to previous test ')
		# print start
		# print end
		for i in range(start+1,end):
			value=self.doc1.Paragraphs(i).Range.Text
			if value[:-1].lower()!="No Low severity faults were observed.".lower():
				
				self.Low.append(value)
		self.Low=filter(lambda name: name.strip(), self.Low)	
		self.faults.append(self.Low)
		
	
		
	def writeinxl(self):
		
		self.xlSht1.Cells(16,4).Value=len(self.High)
		self.xlSht1.Cells(17,4).Value=len(self.Medium)
		self.xlSht1.Cells(18,4).Value=len(self.Low)
		
		i=13
		
		a=["Achilles","Nmap","Nessus","Mu-8000","ISIC"]
		for m in range(0,3):
			for k in range(0,len(self.faults[m])):
				
				self.xlSht1.Rows(i).Insert()
				self.xlSht1.Rows(i).Interior.Color="&hFFFFFF"
				
				self.xlSht1.Cells(i,5).Value=self.faults[m][k]
				s=self.xlSht1.Cells(i,5).Value
				
				for imm in range(0,len(a)):
					if s.find(a[imm]) != -1:
						self.xlSht1.Cells(i,6).Value=a[imm]
				i=i+1
				
		
	def fetchdvcdetails(self):
		rngDoc = self.doc1.Range(0, 0)
		table = self.doc1.Tables(1)
		self.device.append(table.Cell(Row=1, Column=2).Range.Text.split('.')[1])#device name
		self.device.append(table.Cell(Row=1, Column=2).Range.Text.split('.')[2][:-1])#test run
		# print table.Cell(Row=1, Column=2).Range.Text.split('.')[2]
		self.device.append(table.Cell(Row=3, Column=2).Range.Text[:-1])#test date
		table = self.doc1.Tables(4)
		self.device.append(table.Cell(Row=5, Column=2).Range.Text[:-1])#os
		self.device.append(table.Cell(Row=4, Column=2).Range.Text[:-1])#firmware version
		# for i in range(len(self.device)):
			# print self.device[i]
		self.device=filter(lambda name: name.strip(), self.device)
		
	def filldvcdetails(self):
		
		self.xlSht1.Cells(3,2).Value=self.device[0]
		self.xlSht1.Cells(4,2).Value=self.device[3]
		self.xlSht1.Cells(5,2).Value=self.device[4]
		self.xlSht1.Cells(6,2).Value=self.device[1]
		self.xlSht1.Cells(7,2).Value=self.device[2]
		
	def severity(self):
		lenHigh=len(self.High)
		lenMedium=len(self.Medium)
		lenLow=len(self.Low)
		
		i=13
		for ip in range(0,lenHigh):
			self.xlSht1.Cells(i,10).Value="High"
			for m in range(1,11):
				self.xlSht1.Cells(i,m).Font.Color="&hFF"
			i=i+1
		# print i
		for ipp in range(0,lenMedium):
			self.xlSht1.Cells(i,10).Value="Medium"
			for m in range(1,11):
				self.xlSht1.Cells(i,m).Font.Color="&hFF0000"
			i=i+1 
		# print i
		for ippp in range(0,lenLow):
			self.xlSht1.Cells(i,10).Value="Low"
			for m in range(1,11):
				self.xlSht1.Cells(i,m).Font.Color="&h00"
			
			i=i+1
		# print i
	def Srno(self):
		p=13
		total=len(self.High)+len(self.Medium)+len(self.Low)
		for no in range(1,total+1):
			self.xlSht1.Cells(p,1).Value=no
			p=p+1
	def storeforObj(self):
		self.excel2Object.append(excel2(self.doc1,self.xlWb1,len(self.High),len(self.Medium),len(self.Low)))
		for objs in self.excel2Object :
			objs.fetchdvcdetails()
			objs.filldvcdetails()
			objs.NoOfFauts()
			objs.testtype()
			objs.devarrival()
			objs.testedprotocol()
		# self.excel2Object.append(excel2(len(self.High),len(self.Medium),len(self.Low)))
	
	
if __name__ == "__main__":
	r =excel1()
	r.fetchdvcdetails()
	r.filldvcdetails()
	r.highfaults()
	r.mediumfaults()
	r.lowfaults()
	# print r.faults
	r.writeinxl()
	r.severity()
	r.Srno()
	r.storeforObj()
	# print "=============================="
	# print r.faults[2]
	# # print r.High
	# print "=============================="
	# print r.faults[3]
	# # print r.Medium
	# # print "=============================="
	# # print r.Low
	# # print "=============================="
	
	

