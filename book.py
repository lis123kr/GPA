class Book:
	"""
		Contents of a "Excel file" 
	"""

	xls = None

	filename = None

	nsheets = 0

	constraints = 35.0

	sheet_list = []

	# Columns name
	col_DumaPosition = None

	col_DumaSeq = None

	col_Sequence = None

	col_GenomeStructure = None

	col_RepeatRegion = None

	col_ORF = None

	# data list of each columns
	GenomeStructure = []
	
	RepeatRegion = []

	ORF = []

	NCR = []


	# variables for preprocessing & analyzing
	BPRaw = []

	BP = []

	BPxSum = []

	BPxMajor = []
	
	BPxMinor = []

	BPxMAX = []

	P0 = None

	P0_Raw35 = []

	def type_check(self, L):
		for i in range(0,len(L)):
			try:
				L[i] = int(L[i])
			except:
				L[i] = str(L[i])
		return L


	def initialize(self):
		self.filename = self.filename.split('.')[0]

		self.GenomeStructure = self.type_check(self.GenomeStructure)
		self.RepeatRegion = self.type_check(self.RepeatRegion)
		self.ORF = self.type_check(self.ORF)
		self.NCR = self.type_check(self.NCR)






	