from win32com.client import Dispatch 
import os
import win32com.client as win32
from xml.dom import minidom
class Nessus(object):

# This class fetches all data from nessus xml and writes in appropiate positions of Mars.doc

	data = {}
	# version=""
	# feedversion=""
	def __init__(self, path,docfile,xml):
		self.path = path
		word = Dispatch('Word.Application')
		word.Visible=False
		self.xml=xml
		# self.doc1=word.Documents.Open(os.path.join(os.getcwd(), 'MARS.doc'))
		self.doc1=word.Documents.Open(docfile)
		# self.doc1.Save()
		
		# print self.doc1.Tables(6).Columns.Count
		# self.docu = minidom.parse('C:\\users\\inprsha\\desktop\\nessus_report.xml')
		  
		
	def fetchversion(self):
		
	# This function fetches Nessus version and PluginFeed Version from one of the xml file of Nessus
	
		# self.nessusxml = minidom.parse('C:\\users\\inprsha\\desktop\\nessus\\nessus_report_REF620_Dangerous.xml')
		self.nessusxml = minidom.parse(self.xml)
		rprt = self.nessusxml.getElementsByTagName('Report')
		for n in rprt:
			host=n.getElementsByTagName('ReportHost')
			reitem=n.getElementsByTagName('ReportItem')[0]
			self.version=reitem.getElementsByTagName('plugin_output')[0].firstChild.data.split("Nessus version :")[1].split(' ')[1].split('\n')[0]
			self.feedversion= reitem.getElementsByTagName('plugin_output')[0].firstChild.data.split("Plugin feed version :")[1].split(' ')[1].split('\n')[0]
			# print version 
			# print "-------------"
			# print feedversion
	
	def fillversion(self):
	# This function fills the Nessus Version and Plugin Feed Version in MS Word Template.
		table = self.doc1.Tables(3)
		table.Cell(Row=3, Column=2).Range.Text="Nessus Version :" +self.version+"\n"+"Plugin Feed Version:"+self.feedversion
		# table.Cell(Row=3, Column=3).Range.Text="Nessus Version :" +self.version+"\n"+"Plugin Feed Version:"+self.feedversion
		
	def readfromxml(self):
	 
	# This function parses all xml files of Nessus and fetches PluGin Id and their respective analysis,synopsis and Risk Factor.
	
		self.nessusxml = minidom.parse(self.path)
		rprt = self.nessusxml.getElementsByTagName('Report')
		
		
		for n in rprt:
			host=n.getElementsByTagName('ReportHost')
			reitem=n.getElementsByTagName('ReportItem')
			
			for m in reitem:
				
				
				risk =m.getElementsByTagName('risk_factor')[0].firstChild.data
				if risk=='High' or risk=='Medium' or risk=='Critical':
					self.appendToTable(m)
					
					#synopsis =m.getElementsByTagName('synopsis')[0]
					#description=m.getElementsByTagName('description')[0]
					
								
	
	def appendToTable(self, m):
	
	
		#print dir(m)
		self.plugin = {}
		self.pluginID = m.getAttribute('pluginID')
		columns = ['plugin_name', 'synopsis', 'description','risk_factor','fault_details','nessus_analysis']
		for c in columns:
			try:
				self.plugin[c] = m.getElementsByTagName(c)[0].firstChild.data.replace('\n\n', '\n')
			except:
				self.plugin[c] = ''
		self.plugin['ifRealVul'] = self.checkifRealVul('%s (%s)' % (self.plugin['plugin_name'],self.pluginID))
		self.data[self.pluginID] = self.plugin
	
			
	def filltblnessus(self):
	
	#  This function fills Table 2-Summary of Nessus Vulnerabilities.
	
		rngDoc = self.doc1.Range(0, 0)
		# self.doc1.Tables(6).AllowAutoFit =True
		# self.doc1.Tables(6).Columns.Add()
		
		# self.doc1.Tables(6).AutoFitBehavior
		table = self.doc1.Tables(6)
			# table.Cell(Row=1, Column=4).Range.Text="Risk Factor"
		a=2
		for i in self.data:
			
			table.Cell(Row=a, Column=1).Range.Text=i
			table.Cell(Row=a, Column=2).Range.Text=self.data[i]['synopsis']
			table.Cell(Row=a, Column=3).Range.Text=self.data[i]['description']
			table.Cell(Row=a, Column=4).Range.Text=self.data[i]['risk_factor']
			
			if a<=len(self.data):
				a=a+1	
				self.doc1.Tables(6).Rows.Add()
			else:
				# Default PluginIds and their description
				new=self.doc1.Tables(6).Rows.Count
				self.doc1.Tables(6).Rows.Add()
				self.doc1.Tables(6).Rows.Add()
				# print new
				table.Cell(Row=new+1, Column=1).Range.Text="51192" 
				table.Cell(Row=new+2, Column=1).Range.Text="57582"
				table.Cell(Row=new +1, Column=2).Range.Text="The SSL certificate for this service cannot be trusted." 
				table.Cell(Row=new+2, Column=2).Range.Text="The SSL certificate chain for this service ends in an unrecognized self-signed certificate." 
				table.Cell(Row=new+1, Column=3).Range.Text="The server's X.509 certificate does not have a signature from a known public certificate authority. This situation can occur in three different ways, each of which results in a break in the chain below which certificates cannot be trusted. First, the top of the certificate chain sent by the server might not be descended from a known public certificate authority. This can occur either when the top of the chain is an unrecognized, self-signed certificate, or when intermediate certificates are missing that would connect the top of the certificate chain to a known public certificate authority. "
				table.Cell(Row=new+2, Column=3).Range.Text="The X.509 certificate chain for this service is not signed by a recognized certificate authority. If the remote host is a public host in production, this nullifies the use of SSL as anyone could establish a man-in-the-middle attack against the remote host."
		
	def checkifRealVul(self, pluginString):
	
	# This function take input from user and asks whether fetched PlugIn IDs are real vulnerability or not.
	
		print "Does Plugin: \'%s\' a real vulnerability? Y/N" % (pluginString)
		value = raw_input()
		if value == 'Y' or value == 'y':
			return True
		else:
			return False
			
	def fetchpnum(self, searchable, startswith=1):
	# It helps to find the paragraph number where some information is to be appended.
		# print self.doc1.Paragraphs.Count
		for i in range(1, self.doc1.Paragraphs.Count):
			value = self.doc1.Paragraphs(i).Range.Text.lower()
			if value.startswith(searchable.lower()):
				return i
		return 0
	
	def writefaults(self):
	# It fills Appendix E with Plugin Id and its description in accordance to user input.
		# print self.fetchpnum("APPENDIX E: FAULT DETAILS")
		p=len(self.data)
		for i in self.data:
			if self.data[i]['ifRealVul'] == True:
				self.doc1.Paragraphs(self.fetchpnum("APPENDIX E: FAULT DETAILS")).Range.InsertAfter("\n12."+str(p)+" Plugin Id: "+i+"\n"+self.data[i]['synopsis']+"\n")
				p=p-1
				
				
	def searchid(self):
	
	# It writes paragraphs in High Level Overview,Executive Summary,SCADA, Default and Dangerous checks according to PluginId encountered.
	
		# self.doc1.Paragraphs(self.fetchpnum("Comparison with respect to previous test <Only mention the issues, no explanation>")).Range.InsertBefore("Note:/nNessus reported that xxx Service running on the device uses a self-signed SSL certificate. Certificate issued by Certification authority tells the customer that the server information has been verified by a trusted source. Using Self-signed certificate may allow an attacker to launch man in the middle attack. Even though it is not a major security issue, it is recommended to use the certificate signed by a trusted certificate authority. Fault Details This issue is considered as Remark since most of the ABB devices uses Self-Signed SSL Certificate.")
		# self.doc1.Paragraphs(self.fetchpnum("More information about the reported faults can be found in the below subsequent sections .")).Range.InsetBefore("Note:\nXXXX service uses self-signed SSL certificate.")
		self.doc1.Paragraphs(self.fetchpnum("Default Checks")).Range.InsertAfter("SSL Certificate Cannot be Trusted (Plugin ID 51192) plugin reported that SSL certificate of the xxxx service is not signed by a known public authority. The xxx service running on the device uses a self-signed certificate which is not trusted because they are not generated by the trusted Certification Authority. Self-Signed certificates cannot be revoked, which may allow an attacker who has already gained access to monitor and inject data into a connection to spoof an identity if a private key has been compromised. Certifying Authorities on the other hand have the ability to revoke a compromised certificate which prevents it from further use. The device was responding to ICMP ping and TCP connect requests during and after subjecting to this plugin.\nSSL Self-Signed Certificate (Plugin ID 57582) plugin also reported that xxx service uses self-signed SSL Certificate. Issuer name in the SSL Certificate is xxxx This issue is same as one explained above.")
		
		self.doc1.Paragraphs(self.fetchpnum("Dangerous Checks")).Range.InsertAfter("SSL Certificate Cannot be Trusted (Plugin ID 51192) plugin reported that SSL certificate of the xxxx service is not signed by a known public authority. The xxx service running on the device uses a self-signed certificate which is not trusted because they are not generated by the trusted Certification Authority. Self-Signed certificates cannot be revoked, which may allow an attacker who has already gained access to monitor and inject data into a connection to spoof an identity if a private key has been compromised. Certifying Authorities on the other hand have the ability to revoke a compromised certificate which prevents it from further use. The device was responding to ICMP ping and TCP connect requests during and after subjecting to this plugin.\nSSL Self-Signed Certificate (Plugin ID 57582) plugin also reported that xxx service uses self-signed SSL Certificate. Issuer name in the SSL Certificate is xxxx This issue is same as one explained above.")
		
		table = self.doc1.Tables(6)
		# m=self.fetchpnum("TEST DATA")
		# print m
		for i in range(1,self.doc1.Tables(6).Rows.Count):
			
			if table.Cell(i,1).Range.Text[:-2]=='23812':
				self.doc1.Paragraphs(self.fetchpnum("More information about the reported faults can be found in the below subsequent sections.")).Range.InsertAfter("The device has an ICCP/COTP TSAP Addressing Weakness.\n".capitalize())
				self.doc1.Paragraphs(self.fetchpnum("For all other tests no anomaly was reported in the <Device Name>.")).Range.InsertAfter("ICCP/COTP TSAP Addressing Weakness (Plugin ID 23812) plugin reported that it is possible to determine a COTP TSAP value on the remote ICCP server by trying possible values. The ICCP stack (protocols MMS and IEC 61850) includes ISO 7073 at the Transport Layer. ISO 7073 specifies the Connection Oriented Transport Protocol (COTP) that includes a pair of user configurable 16-bit numeric, or in some cases ASCII string values, to identify client endpoints called Transport Service Access Points (TSAP's). The TSAP used in the host server was guessed by trying a sample of possible values that are commonly used and easily attempted by trial-and-error. It is recommended that the Transport Service Access Points values be randomized, or (if possible) Secure ICCP be used instead.\n ".capitalize())
				self.doc1.Paragraphs(self.fetchpnum("SCADA Checks")).Range.InsertAfter("ICCP/COTP TSAP Addressing Weakness (Plugin ID #23812) plugin reported that it is possible to determine a COTP TSAP value on the remote ICCP server by trying possible values. The ICCP stack (and other protocols MMS and IEC 61850) includes ISO 7073 at the Transport Layer. ISO 7073 specifies the Connection Oriented Transport Protocol (COTP) that includes a pair of user configurable 16-bit numeric, or in some cases ASCII string values, to identify client endpoints called Transport Service Access Points (TSAP's). The TSAP used in the ho host server was guessed by trying a sample of possible values that are commonly used and easily attempted by trial-and-error.\n".capitalize())
			
			if table.Cell(i,1).Range.Text[:-2]=='23811':
				self.doc1.Paragraphs(self.fetchpnum("More information about the reported faults can be found in the below subsequent sections.")).Range.InsertAfter("COTP (ISO 7073) is running on the host and may be part of an ICCP server, MMS application, or substation automation device that uses IEC61850-UCA \n".capitalize())
				self.doc1.Paragraphs(self.fetchpnum("For all other tests no anomaly was reported in the <Device Name>.")).Range.InsertAfter("ICCP/COTP (ISO 7073) Protocol Detection (Plugin ID 23811) plugin reported that COTP protocol running on the device which is a part of MMS application is an unprotected binary protocol. Manufacturing Message Specification (MMS) deals with messaging system for transferring real time process data and supervisory control information between networked devices and/or computer applications. Since the communication between the MMS client and the server is not encrypted, it allows intruders to gain access to the control system data. There are many more vulnerabilities reported with respect to this protocol like bypassing intended controls, integrity violation, eavesdropping by non-trusted entities, spoofing and Playback of captured data from non-trusted entities etc. It is recommended that access to the ICCP server be limited, or (if possible) secure ICCP be used instead. \n".capitalize())
				self.doc1.Paragraphs(self.fetchpnum("SCADA Checks")).Range.InsertAfter("ICCP/COTP (ISO 7073) Protocol Detection (Plugin ID #23811) plugin reported that COTP protocol running on the device which is a part of MMS application is an unprotected binary protocol. Manufacturing Message Specification (MMS) deals with messaging system for transferring real time process data and supervisory control information between networked devices and/or computer applications. Since the communication between the MMS client and the server is not encrypted, it allows intruders to gain access to the control system data. There are many more vulnerabilities reported with respect to this protocol like bypassing intended controls, integrity violation, eavesdropping by non-trusted entities, spoofing and Playback of captured data from non-trusted entities etc. Secure ICCP afford the most prudent and secure communications between control and reliability center exchanging data over network.\n".capitalize())
			
			if table.Cell(i,1).Range.Text[:-2]=='12213':
				
				self.doc1.Paragraphs(self.fetchpnum("More information about the reported faults can be found in the below subsequent sections.")).Range.InsertAfter("XXXX service uses self-signed SSL certificate.")
				self.doc1.Paragraphs(self.fetchpnum("For all other tests no anomaly was reported in the <Device Name>.")).Range.InsertAfter("Nessus reported that xxx Service running on the device uses a self-signed SSL certificate. Certificate issued by Certification authority tells the customer that the server information has been verified by a trusted source. Using Self-signed certificate may allow an attacker to launch man in the middle attack. Even though it is not a major security issue, it is recommended to use the certificate signed by a trusted certificate authority. Fault Details This issue is considered as Remark since most of the ABB devices uses Self-Signed SSL Certificate.")
				self.doc1.Paragraphs(self.fetchpnum("Default Checks")).Range.InsertAfter("TCP/IP Sequence Prediction Blind Reset Spoofing DoS (Plugin ID#12213) This plugin reported that it may be possible to send spoofed RST packets to the remote system. This plugin creates a TCP session to an open port on the device. It sends a character to the port, if it gets a reply other than RST packet it indicates that the session is alive, and then it spoofs a RST with the sequence number incremented by 512 from the valid tuple defining the socket (i.e. srchost, dsthost, srcport, dstport). It then sends a character to the socket created to check for a RST from the host. If it gets a RST from the device, then that indicates that the system accepted and processed the spoofed RST. When this scan was launched on IED650, from the packet capture it was observed that the spoofed RST packet closed the TCP session successfully. This is a known issue in most of the operating systems, hence it is considered as Remark\n".capitalize())
				
			
				
			
if __name__ == "__main__":
	p = Nessus()
	
	p.fetchversion()
	p.fillversion()
	p.readfromxml()

	
	p.filltblnessus()
	
	p.searchid()
	print "Please wait... "
	p.writefaults()
	
	#print p.ver			

	
	
	
	