from win32com.client import Dispatch 
import os
import win32com.client as win32
class robustnessform(object):

# This class fetches all data from robustness form which has to be copied in MS Word Template

	data = {}
	def __init__(self,path,docfile):
		word = Dispatch('Word.Application')
		word.Visible=True
		self.path=path
		self.doc = word.Documents.Open(os.path.join(os.getcwd(), self.path))
		self.doc1=word.Documents.Open(docfile)
		# self.doc1=word.Documents.Open(os.path.join(os.getcwd(), 'MARS.doc'))
		self.doc1.Save()
		
	
				
	def fetchtabledata(self):
	
		# This function is used for fetching  
								# (i)Table 1 -Product Details from Robustness Form and storing in a dictionary 'data'
								# (ii)Table 7 - Primary contact person and storing in a dictionary 'data2'
	
		rngDoc = self.doc.Range(0, 0)
		table = self.doc.Tables(1)
		self.data = {}
		for i in range(2,6):
			key=table.Cell(Row=i, Column=1).Range.Text[:-2].lower()
			value=table.Cell(Row=i, Column=2).Range.Text[:-2]
			self.data[key]=value
	
		
		table = self.doc.Tables(15)
		self.data2 = {}
		for i in range(2,6):
			key=table.Cell(Row=i, Column=1).Range.Text[:-2].lower()
			value=table.Cell(Row=i, Column=2).Range.Text[:-2]
			self.data2[key]=value
			
	def writetabledata(self):
	
	# This function is used to fill following from data and data2 dictionary in MS Word Template
								  #  i)3.4	Device under test (DUT ) 
								  # ii)3.2	Product Contact Person 
	
		rngDoc = self.doc1.Range(0, 0)
		table = self.doc1.Tables(4)
		
		table.Cell(Row=2, Column=2).Range.Text=self.data['product name']
		table.Cell(Row=4, Column=2).Range.Text=self.data['firmware version']
		table.Cell(Row=5, Column=2).Range.Text=self.data['operating system ']	
			
			
		table = self.doc1.Tables(2)
		table.Cell(Row=2, Column=2).Range.Text=self.data2['name']
		table.Cell(Row=3, Column=2).Range.Text=self.data2['location']
		table.Cell(Row=4, Column=2).Range.Text=self.data2['email']
		table.Cell(Row=5, Column=2).Range.Text=self.data2['telephone']	
			
	def replaceDeviceName(self):
	
		# This function is used to replace <Device Name> in Paragraphs and Headers by actual Device Name which is stored in dictionary 'data'
		
		##### Paragraphs ######
		for i in range(1, self.doc1.Paragraphs.Count):
		
			# print "hello"
			para=self.doc1.Paragraphs(i).Range.Text.lower()
			find=para.find("<Device Name>".lower())
			if find > 0 :
				# print str(find) + "in para" + str(i)
				# print type(self.doc1.Paragraphs(i).Range.Text)
				para=para.replace("<Device Name>".lower(),self.data['product name'])
				self.doc1.Paragraphs(i).Range.Text=para.title()
		
		
		
		##### Headers ######
		text = self.doc1.Sections(1).Headers(win32.constants.wdHeaderFooterPrimary).Range.Tables(1)
		text.Cell(3,1).Range.Text = "Test Report for "+self.data['product name']
		# print text.Cell(3,1).Range.Text
		text1 = self.doc1.Sections(1).Headers(win32.constants.wdHeaderFooterFirstPage).Range.Tables(1)
		text1.Cell(4,2).Range.Text = "Test Report for " +self.data['product name']
		# print text1.Cell(4,2).Range.Text
		


if __name__ == "__main__":
	r = robustnessform()
	r.fetchtabledata()
	r.writetabledata()
	r.replaceDeviceName()
	
	# print r.data2
	
	
	

