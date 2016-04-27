from win32com.client import Dispatch 
from numpy import *
import os
import win32com.client as win32
from xml.dom import minidom
class AchillesFuzz(object):
# This class fetches all data from achilles xml and writes in appropiate positions of  MS Word Template
	data={}
	tstname2=list()
	rslt=[]
	rslt2=[]
	
	def __init__(self,path,docfile):
		word = Dispatch('Word.Application')
		word.Visible=True
		# self.doc1=word.Documents.Open(os.path.join(os.getcwd(), 'MARS.doc'))
		self.doc1=word.Documents.Open(docfile)
		self.path=path
		self.xmlfile = minidom.parse(self.path)
		# self.table = self.doc1.Tables(11)
		self.table2 = self.doc1.Tables(12)
		self.doc1.Save()
		
	def fetchfuzztestname(self):
	
	
	# Fetching  SubTest Names from 2nd column of Table 8 - Achilles Level2 Fuzz Test Results of MS Word Template about which data is to be searched in xml file of Achilles
	
		for i in range(33) :
			i=i+1
			a = self.table2.Cell(Row=i, Column=2).Range.Text[:-2]
			if i==2 or i==11 or i==13 or i==18 or i==33 or i==30 or i==29 : 
				a +=" (L1/L2)" 
				self.tstname2.append(a)
			elif i==1:
				print "  "
			else :
				a +=" (L2)"  
				self.tstname2.append(a)		
	
	
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
		
	
	def fetchfuzz(self):
	
	# merging data with their respective Subtest Names
	
		self.rslt2=self.fetchxml(self.tstname2)
		# print self.rslt2

	def fillfuzz(self):
	
	# Filling data in Table 8 of MS Word Template
	
		for rslt in self.rslt2:
			for subtest in range(2,34) :
				self.table2.Cell(Row=subtest, Column=4).Range.Text="Normal"
				a = self.table2.Cell(Row=subtest, Column=2).Range.Text[:-2]
				# print a
				if subtest==2 or subtest==11 or subtest==13 or subtest==18 or subtest==33 or subtest==30  or subtest==29  :  
					a +=" (L1/L2)"  
				else :
					a +=" (L2)"
				
				if rslt[0]==a:
					f=self.table2.Cell(Row=subtest, Column=7).Range.Text
					f=f+rslt[1].split(" (")[0]+":"+rslt[2]
					self.table2.Cell(Row=subtest, Column=7).Range.Text=f
					status=self.table2.Cell(Row=subtest, Column=5).Range.Text
					if status=="Normal" and rslt[2]!="Normal":
						status=rslt[2]
					elif status=="Warning" and rslt[2]!="Failure":
						status="Warning"
					elif status=="Failure":
						status="Failure"
					else:
						status=rslt[2]
					self.table2.Cell(Row=subtest, Column=5).Range.Text=status
		
	
					

			
	
if __name__ == "__main__":
	p = AchillesFuzz()
	

	
	
	p.fetchfuzztestname()
	p.fetchfuzz()
	p.fillfuzz()
	
