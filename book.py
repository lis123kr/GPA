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

	col_basepair = ["A","G","C","T"]
	# data list of each columns
	GenomeStructure = []
	
	RepeatRegion = []

	ORF = []

	NCR = []

	# variables for preprocessing & analyzing
	BPRaw = []

	BPRawLength = []

	BP35 = []

	BPxMajor = []
	
	BPxMinor = []

	P0 = None

	Dumas = None
	
	def type_check(self, L):
		for i in range(0,len(L)):
			try:
				L[i] = int(L[i])
			except:
				L[i] = str(L[i])
		return L

	def initialize(self):
		self.filename = self.filename.split('.')[0]
		self.nsheets = len(self.sheet_list)
		self.GenomeStructure = self.type_check(self.GenomeStructure)
		self.RepeatRegion = self.type_check(self.RepeatRegion)
		self.ORF = self.type_check(self.ORF)
		self.NCR = self.type_check(self.NCR)

	def mergesheets(self):
		from pandas import merge

		for x in self.sheet_list:
			sheet = self.xls.parse(x)
			if self.P0 is None:
				self.P0 = sheet
			else:
				self.P0 = merge(self.P0, sheet, how='outer', on=[self.col_GenomeStructure, 
					self.col_RepeatRegion, self.col_ORF, self.col_DumaPosition, self.col_DumaSeq])
			sheet = sheet[ sheet[self.col_Sequence] != '-' ]
			self.BPRaw.append(sheet)
			self.BPRawLength.append(len(sheet))
		
		self.Dumas = self.P0[ self.P0[self.col_DumaSeq] != '-' ]
		
	def get_major_minor(self, Passage):	
		from numpy import zeros
		from pandas import DataFrame
		"""	
			- major : 염기 A, G, C, T 중 가장 큰 값
			- minor : 염기 중 두번째로 큰 값, argsort 후 인덱스 2의 값을 추출
		"""
		ranked_df = Passage.values.argsort(axis=1)		
		arr = zeros((len(Passage),1))
		for i, d in enumerate(ranked_df[:,2]):
			arr[i] = Passage[i:i+1][self.col_basepair[d]]
		major = DataFrame(Passage.max(axis=1))
		minor = DataFrame(arr, index=major.index)		
		return major, minor
