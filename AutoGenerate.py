import os
import win32com.client as win32
import datetime
from achillesstorm import AchillesStorm
from achillesfuzz import AchillesFuzz
from nmap import Nmap
from robust import robustnessform


class autogenerate:

# This class intergrates working of nmap.py,achillesstorm.py,achillesfuzz.py and robust.py.This should be run after FileManagerNessus.py

	NmapObject=[]
	RobustObject=[]
	AchStormObj=[]
	AchFuzzObj=[]
	
	def __init__(self):
		self.root = '.'
		self.fileList = {}
		
	def makeRoot(self, root):
		self.root = root
		
	def isXML(self, fileName):
	
	# This function finds for the xml files in the root folder provided.
	
		if fileName.endswith('.xml') or fileName.endswith('.doc') and fileName.lower().rfind('regression')==-1  :
			return True
		else:
			return False
			
	def makeFileList(self):
	
	# This function create list of file names and file paths  having xml extension .
	
		key=0
		for (path, dirs, files) in os.walk(self.root):
			for item in files:
				fileName = os.path.join(path,item)
				
				if self.isXML(fileName):
					
					value=(fileName)
					key=key+1
					self.fileList[key]=value
		
		# for i in self.fileList :
			
			# print i
			# print self.fileList[i]
			
			
	def selectFile(self):
	
	# This function reads first few lines of every xml file encounterd in the root to recognize which xml file is meant for which Tool.
	
		for i in self.fileList:
			
			file=open(self.fileList[i])
			xml=file.read()
			if "Robustness".lower() in self.fileList[i].lower() :
				
				self.RobustObject.append(robustnessform(self.fileList[i],docfile))
				
			elif "Achilles" in xml:
				# print i,'--->',self.fileList[i]
				# print "achu"
				self.AchStormObj.append(AchillesStorm(self.fileList[i],docfile))
				self.AchFuzzObj.append(AchillesFuzz(self.fileList[i],docfile))
				
			elif "nmaprun" in xml :
				
				self.NmapObject.append(Nmap(self.fileList[i],docfile))
				
	def storeforObj(self):
	
	# This functions calls for the required functions for every test tool object.
	
		for objs in self.RobustObject:
			# print objs.data.keys()
			
			objs.fetchtabledata()
			objs.writetabledata()
			objs.replaceDeviceName()
		print "-----Robust done-----------"
		
		for objs in self.NmapObject:
			objs.fetchip()
			objs.fillip()
			objs.fetchmac()
			objs.fetchversion()
			objs.fillversion()
			objs.fillmac()
			# objs.cleartblnmap()
			objs.fetchtblfromxml()								
			objs.filltblnmap()
			
			
		print "---------Device Profiling Tool done-----------"		
		
		for objs in self.AchStormObj:
			
			objs.clearncorrect()
			objs.fetchnfillIPStack()
			objs.fetchports()
			objs.fillports()
			objs.fetchstormtestname()
			objs.fetchstorm()
			objs.fillstorm()
		print "---------------Fuzzing and Flooding Tool Storm Results Done ---------------"	
		
		for objs in self.AchFuzzObj:	
			objs.fetchfuzztestname()
			objs.fetchfuzz()
			objs.fillfuzz()
		print "---------Fuzzing and Flooding Tool Fuzz Results  Done-------------------"	
			
		
		
		
		
if __name__ == "__main__":
	fp = autogenerate()
	print "Enter path for Root folder of xml files"
	root =raw_input()
	# root = "C:\Users\inprsha\Desktop\\"
	print "Enter path for REPORT TEMPLATE"
	docfile=raw_input()
	# docfile= "C:\Users\inprsha\Desktop\Mars.doc
	fp.makeRoot('.')
	fp.makeFileList()
	fp.selectFile()
	fp.storeforObj()
