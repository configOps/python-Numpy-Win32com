import os
from nessus import Nessus
from xml.dom import minidom

class FilemangerNessus(object):

# This class manages all xml files of Nessus .

	NessusObject=[]
	ObjDict={}
	def __init__(self, path):
		self.path = path
		
		                   
	def listFiles(self):
	# This function is used to list files present in the provided folder and create objects for every file.
		for f in os.listdir(self.path):
			print "Reading File %s" % f
			self.pth=os.path.join(self.path, f)
			self.NessusObject.append(Nessus(self.pth,docfile,xml))
	
	def storeforObj(self):
	# This fucntion calls required functions for all objects.
		for objs in self.NessusObject:
			# print objs.data.keys()
			objs.fetchversion()
			objs.fillversion()
			objs.readfromxml()
			self.ObjDict.update(objs.data)
		
		objs.filltblnessus()
		objs.searchid()
		objs.writefaults()
		print "----------------------------"
		# print self.ObjDict

if __name__ == "__main__":
	print "Give path for any of the xml file of Nessus to fetch Nessus and PluginFeed Version e.g" 
	print "C:\\users\\inprsha\\desktop\\nessus\\nessus_report_REF620_Dangerous.xml"
	xml=raw_input()
	i=0
	m=""
	p=len(xml.split("\\"))-1
	while(i<p):
		m=m+xml.split("\\")[i] +"\\"
		i=i+1
	print m
	p = FilemangerNessus(m)
	print "Enter path for MS Word Report Template"
	docfile=raw_input()
	p.listFiles()
	p.storeforObj()
	