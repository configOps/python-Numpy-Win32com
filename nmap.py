from win32com.client import Dispatch 
import os
import win32com.client as win32
from xml.dom import minidom
class Nmap(object):

# This class fetches all data from nmap xml and writes in appropiate positions of MS Word Report Template.

	
	def __init__(self,path,docfile):
	
		word = Dispatch('Word.Application')
		word.Visible=True
		self.doc1=word.Documents.Open(docfile)
		#os.path.join(os.getcwd(), 
		self.path=path
		# print self.path
		self.xmlfile = minidom.parse(self.path)
		self.dut = {}
		self.nmaptbldata = {}
		self.doc1.Save() 
		
	def fetchip(self):
	
	# fetching Ip address from xml file of Nmap 
	
		self.dut={}
		nok = self.xmlfile.getElementsByTagName('address')
		for n in nok:
			type=n.getAttribute('addrtype')
			name =n.getAttribute('addr')
			if type=="ipv4" :
				# print name
				name=str(name)
				break
		rngDoc = self.doc1.Range(0, 0)
		key='1'
		value=name
		self.dut[key]=value
		
	def fetchmac(self):
	
	# fetching Mac address from xml file of Nmap 	
		nok = self.xmlfile.getElementsByTagName('address')
		for n in nok:
			type=n.getAttribute('addrtype')
			name =n.getAttribute('addr')
			if type=="mac" :
				name=str(name)
				break
			rngDoc = self.doc1.Range(0, 0)
			key='2'
			value=name
			self.dut[key]=value	
			
	def fillip(self):
	
	#filling "DUT interface connected to a Mu-4000 port" in section 8.1.1 to 8.1.4 of MS Word Template
	
		for i in range(14,18):
			
			table = self.doc1.Tables(i)
			table.Cell(Row=3, Column=2).Range.Text=self.dut['1']
			
	def fillmac(self):
	
	# Filling mac address in 3.4 Device Under Test of MS Word Template
	
		table = self.doc1.Tables(4)
		table.Cell(Row=3, Column=2).Range.Text=self.dut['2']
			
		
	def fetchversion(self):
	
	# Fetching Nmap Version from xml file of Nmap
		
		scan=self.xmlfile.getElementsByTagName('nmaprun')[0]
		key='3'
		value=scan.getAttribute('version')
		self.dut[key]=value
		
	def fillversion(self):
	
	# Filling Nmap Version in 3.3 Test Equipment of MS Word Template
		table = self.doc1.Tables(3)
		table.Cell(Row=4, Column=2).Range.Text=self.dut['3']
		# table.Cell(Row=4, Column=3).Range.Text=self.dut['3']
		table.Cell(Row=5, Column=3).Range.Text="0.07"	
		table.Cell(Row=5, Column=2).Range.Text="0.07"	
		
		
	def fetchtblfromxml(self):
	
	# Fetching nmap Ports and their description from xml file of Nmap
	
		nrun = self.xmlfile.getElementsByTagName('nmaprun')[0]
		host = nrun.getElementsByTagName('host')[0]
		ports = host.getElementsByTagName('ports')[0]
		port = ports.getElementsByTagName('port')
		i=2
		for f in port:
		
			self.doc1.Tables(5).Rows.Add()
			key=f.getAttribute('portid')+"/"+f.getAttribute('protocol')
			state=f.getElementsByTagName('state')[0]
			try:
				service=f.getElementsByTagName('service')[0]
				# print service.getAttribute('name')
				# table.Cell(Row=i, Column=3).Range.Text=service.getAttribute('name')
				value=[state.getAttribute('state'),service.getAttribute('name')]
				self.nmaptbldata[key]=value
			except:
				value=[state.getAttribute('state'), ' ']
				self.nmaptbldata[key]=value
				
			# print "-------------------------------------------------------------------------------------------------------"
			i=i+1
			
	def filltblnmap(self):
	
	# Filling Table 1 Nmap Test Results in  MS Word Template
	
		rngDoc = self.doc1.Range(0, 0)
		table = self.doc1.Tables(5)
	
		a=2
		for i in self.nmaptbldata:
			
			table.Cell(Row=a, Column=1).Range.Text= i
			table.Cell(Row=a, Column=2).Range.Text=self.nmaptbldata[i][0]
			try:
				table.Cell(Row=a, Column=3).Range.Text=self.nmaptbldata[i][1]
			except:
				table.Cell(Row=a, Column=3).Range.Text=" "
			if a<=len(self.nmaptbldata):
				a=a+1	
			else:
				break
		self.doc1.Save()
		
	# def fetchtblnmap(self):
	
		# # This function is for fetching existing data in nmap results table of Mars.doc
	
		# rngDoc = self.doc1.Range(0, 0)
		# table = self.doc1.Tables(5)
		# self.exsdata = {}
		# for a in range(2,self.doc1.Tables(5).Rows.Count):
			# key=table.Cell(Row=a, Column=1).Range.Text[:-2]
			# #srvName = table.Cell(Row=a, Column=3).Range.Text[:-2]
			# #if srvName != ' ':
			# value=[table.Cell(Row=a, Column=2).Range.Text[:-2],table.Cell(Row=a, Column=3).Range.Text[:-2]]
			# #else:
			# #	value=[table.Cell(Row=a, Column=2).Range.Text[:-2]]
			# self.exsdata[key]=value
	
	# def cleartblnmap(self):	
		
		# rngDoc = self.doc1.Range(0, 0)
		# table = self.doc1.Tables(5)
		# count = self.doc1.Tables(5).Rows.Count
		# # print count
		# for i in range(2,count):
			# ##everytime we are deleting 2nd row and by doing so rows are shifting up
			# self.doc1.Tables(5).Rows(2).Delete()
			
	# def cmpnmapdata(self):
	
		# nmaptbl =set(self.nmaptbldata.keys())
		# exsdat=set(self.exsdata.keys())
		# # print nmaptbl.difference(exsdat)
		# # print "**************************"
		# # print cmp(self.nmaptbldata,self.exsdata)
		# # print "**************************"
		
if __name__ == "__main__":
	p = Nmap()
	p.fetchip()
	p.fillip()
	p.fetchmac()
	p.fetchversion()
	p.fillversion()
	p.fillmac()
	# p.cleartblnmap()
	p.fetchtblfromxml()								
	p.filltblnmap()
	# p.fetchtblnmap()
	# p.cmpnmapdata()
	
	
	


	
	
