from win32com.client import Dispatch 
from numpy import *
import os
import win32com.client as win32
from xml.dom import minidom
class AchillesStorm(object):
# This class fetches all data from achilles xml and writes in appropiate positions of Mars.doc
	data={}
	tstname1=list()
	rslt=[]
	rslt1=[]
	
	
	def __init__(self,path,docfile):
		word = Dispatch('Word.Application')
		word.Visible=True
		
		self.doc1=word.Documents.Open(docfile)
		# self.doc1=word.Documents.Open(os.path.join(os.getcwd(), 'MARS.doc'))
		self.path=path
		self.xmlfile = minidom.parse(self.path)
		self.table = self.doc1.Tables(11)
		self.table2 = self.doc1.Tables(12)
		self.doc1.Save()
		
		
	def fetchpnum(self, searchable, startswith=1):
		# print self.doc1.Paragraphs.Count
		for i in range(1, self.doc1.Paragraphs.Count):
			value = self.doc1.Paragraphs(i).Range.Text.lower()
			if value.startswith(searchable.lower()):
				return i
		return 0
		
	def fetchnfillIPStack (self):
		# Fetching speed and mode from xml and writing to 4.3.1 section of MS Word Template
		
		lmod = self.xmlfile.getElementsByTagName('link-modes')
		for nn in lmod:
			port=nn.getElementsByTagName('port')[0]
			for mo in lmod:
				
				mode =mo.getElementsByTagName('mode')[0]
				abr=mode.firstChild.data.split('-')[0]
				abbr=mode.firstChild.data.split('- ')[1]
		
		self.doc1.Paragraphs(self.fetchpnum("Table 3 shows the various protocol tests and results")).Range.Text="Table 3 shows the various protocol tests and results. The Ethernet link is set to " +abr+  abbr+ ".\n"
		
	
	
	def fetchports(self):
	
	# FETCHING PORT NUMBERS FROM XML FILES OF Achilles
	
		nodes = self.xmlfile.getElementsByTagName('testcase-report')
		for node in nodes:
			config=node.getElementsByTagName('configuration')
			for n in config:
				mp=n.getElementsByTagName('parameter')  	
				for m in mp:
					name =m.getElementsByTagName('name')[0]
					value=m.getElementsByTagName('value')[0]
					if name.firstChild.data=="Destination TCP Ports" :
						openports=value.firstChild.data.split('Open ports:')[1].split(";")[0]
						key="1"
						value=openports
						self.data[key]=value
					
					if name.firstChild.data=="Destination UDP Ports" :
						open=value.firstChild.data.split('Open ports:')[1].split(";")[0]	
						key="2"
						value=open
						self.data[key]=value
					break
					
				break
				
	def fillports(self):
	
	# Filling port numbers in TABLE 7 AND 8 of  MS Word Template
	
		for i in range(14,22):
			self.table.Cell(Row=i, Column=3).Range.Text=self.data["1"]
		for ip in range(18,30):
			self.table2.Cell(Row=ip, Column=3).Range.Text=self.data["1"]
		for i in range(22,25):
			self.table.Cell(Row=i, Column=3).Range.Text=self.data["2"]
		for ip in range(30,34):
			self.table2.Cell(Row=ip, Column=3).Range.Text=self.data["2"]
			
	def fetchstormtestname(self):
	
	# Fetching  Test Names from 2nd column of Table 7 - Achilles Level2 Storm Test Results of MS Word Template about which data is to be searched in xml file of Achilles
	
	
		for i in range(24) :
			i=i+1
			a = self.table.Cell(Row=i, Column=2).Range.Text[:-2]
			if i==13 or i==15 or i==17 or i==18 or i==19 or i==20 or i==21 : 
				a += " (L2)"  
				self.tstname1.append(a)
			elif i==1:
				print " "
			else :
				a += " (L1/L2)"  
				self.tstname1.append(a)
				

	
	def fetchxml(self,tstname):
	
	# Fetching data for every SubTest from  xml file of Achilles.
		
		trprt = self.xmlfile.getElementsByTagName('testcase-report')
		i=0;
		l=0
		f=0;
		flag=0;
		for n in trprt:

			nod=n.getElementsByTagName('name')[0]
			for index in range(len(tstname)):
				tst=tstname[index] #p="Ethernet Unicast Storm (L1/L2)"
				if nod.firstChild.data==tst:
					sum=n.getElementsByTagName('summary')[0]
					test=sum.getElementsByTagName('test')[0]
					mp=test.getElementsByTagName('monitor')  	
					f=f+1
					for m in mp:
						if m.parentNode.nodeName=='test':
							if flag==0:
								i=i+1
							name =m.getElementsByTagName('name')[0]
							value=m.getElementsByTagName('value')[0]  
							data=name.firstChild.data
							color=value.firstChild.data.split('_')[1]
							if color=='red':
								status='Failure'
							elif color=='yellow':
								status='Warning'
							elif color=='green':
								status='Normal'
							if f==1:
								self.rslt.append([])
								self.rslt[i-1].append(tst)
								self.rslt[i-1].append(data)
								self.rslt[i-1].append(status)
							else:  #print i
								flag=0
								for x in range(0,i-1): #print x;
									if self.rslt[x][0]==tst: #print "here1"
										if self.rslt[x][1]==data: #print "heere2",k
											flag=1
											if self.rslt[x][2]=='Normal' and status!='Normal':
												self.rslt[x][2]=status
											
											
											elif self.rslt[x][2]=='Warning' and status!='Failure':
												self.rslt[x][2]='Warning'
											
											
											elif self.rslt[x][2]=='Failure':
												self.rslt[x][2]=='Failure'
											
								if flag==0:
									self.rslt.append([])
									self.rslt[i-1].append(tst)
									self.rslt[i-1].append(data)
									self.rslt[i-1].append(status)
									flag=0
									
		return self.rslt
		
	def fetchstorm(self):
	
	# merging data with their respective test Name
	
		self.rslt1=self.fetchxml(self.tstname1)
		# print self.rslt1
		# print type(self.rslt1[0])
		# print self.rslt1[0]
		# print "----------------------------------"
	
	

	def fillstorm(self):
	
	# Filling data in Table 7 of MS Word Template
	
		for rslt in self.rslt1:
			for testname in range(2,25) :
				self.table.Cell(Row=testname, Column=4).Range.Text="Normal"
				a = self.table.Cell(Row=testname, Column=2).Range.Text[:-2]
				# Appending (L2) or (L1/L2) as per xml file naming convention
				if testname==13 or testname==15 or testname==17 or testname==18 or testname==19 or testname==20 or testname==21 : 
					a += " (L2)"  
				else :
					a += " (L1/L2)" 
				
				if rslt[0]==a:
					f=self.table.Cell(Row=testname, Column=6).Range.Text
					f=f+rslt[1].split(" (")[0]+":"+rslt[2]
					self.table.Cell(Row=testname, Column=6).Range.Text=f
					status=self.table.Cell(Row=testname, Column=5).Range.Text
					if status=="Normal" and rslt[2]!="Normal":
						status=rslt[2]
					elif status=="Warning" and rslt[2]!="Failure":
						status="Warning"
					elif status=="Failure":
						status="Failure"
					else:
						status=rslt[2]
					self.table.Cell(Row=testname, Column=5).Range.Text=status
		
	
					
	def clearncorrect(self):
	
# Correcting spellings of  test name and subtest name in MS Word Template
# Clearing default comments in MS Word Template	
		self.table.Cell(Row=2, Column=6).Range.Text="  "
		self.table2.Cell(Row=2, Column=7).Range.Text="  "
		self.table2.Cell(Row=2, Column=5).Range.Text="  "
		self.table.Cell(Row=2, Column=5).Range.Text="  "
		self.table2.Cell(Row=7, Column=2).Range.Text="Ethernet VLAN/LLC/SNAP Chaining"
		self.table2.Cell(Row=5, Column=2).Range.Text="Ethernet LLC/SNAP Grammar"
		self.table.Cell(Row=7, Column=2).Range.Text="ARP Cache Saturation Storm"
		for i in range(2,25) :
			self.table.Cell(Row=i, Column=4).Range.Text="Normal"	
		for i in range(2,34):
			self.table2.Cell(Row=i, Column=4).Range.Text="Normal"	
			
	
if __name__ == "__main__":
	p = AchillesStorm()
	
	p.clearncorrect()
	p.fetchnfillIPStack()
	p.fetchports()
	p.fillports()
	p.fetchstormtestname()
	p.fetchstorm()
	p.fillstorm()
	
	
	
	
	
