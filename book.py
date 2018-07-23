class Book:
	"""
		Contents of a "Excel file" 
	"""
	def __init__(self):
		self.xls = None

		self.filename = None

		self.nsheets = 0

		self.constraints = 35.0

		self.sheet_list = []

		self.sheets = []
		# Columns name
		self.col_DumaPosition = None

		self.col_DumaSeq = None

		self.col_Sequence = None

		self.col_GenomeStructure = None

		self.col_RepeatRegion = None

		self.col_ORF = None

		self.col_Position = None

		self.col_basepair = ["A","G","C","T"]
		# data list of each columns
		self.GenomeStructure = []
		
		self.RepeatRegion = []

		self.ORF = []

		self.NCR = []

		# variables for preprocessing & analyzing
		self.BPRaw = []

		self.BPRawLength = []

		self.BP35 = []

		self.BPxMajor = []
		
		self.BPxMinor = []

		self.P0 = None

		self.Dumas = None
	
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

	def preprocessing(self):
		from analyzer import infolog
		from pandas import DataFrame
		infolog("preprocessing start")
		for bp in self.BPRaw:
			self.BPRawLength.append(len(bp))

			# 염기 A, G, C, T의 합계 컬럼 추가
			bp["sum"] = bp[self.col_basepair].sum(axis=1)

			# 35bp
			bp = bp[[self.col_GenomeStructure, self.col_RepeatRegion, self.col_DumaPosition,
					self.col_DumaSeq, self.col_ORF, self.col_Sequence, "sum"] + self.col_basepair]
			# sum 값이 35 이상인 데이터 추출
			# constraions = 35.0
			bp = bp[ bp["sum"] >= self.constraints ]
			pmajor, pminor = self.get_major_minor(bp[self.col_basepair])
			# major, minor의 값을 추출
			self.BPxMajor.append(pmajor)
			self.BPxMinor.append(pminor)
			# pminor, pmajor가 값을 가지지 않을 때 에러 발생
			bp['minor'] = pminor if len(pminor) is not 0 else 0
			bp['major'] = pmajor if len(pmajor) is not 0 else 0
			
			self.BP35.append(bp)			

		# major, minor의 인덱스 컬럼 추가
		for bp in self.BP35:
			argx = -bp[self.col_basepair] 	# A, G, C, T의 순으로 우선순위를 위해 음수로 정렬
			bps = argx.values.argsort(axis=1)
			bp['minor_idx'] = DataFrame(bps[:,1], index=bp.index)
			bp['major_idx'] = DataFrame(bps[:,0], index=bp.index)
		infolog("preprocessing end")

	def mergesequence(self):
		from analyzer import infolog
		from pandas import merge
		infolog("mergesequence start")
		# print(type(self.sheet_list), len(self.sheet_list))
		for sheet in self.sheets:
			# sheet = self.xls.parse(x)
			# print(type(sheet))
			if self.P0 is None:
				self.P0 = sheet
			else:
				self.P0 = merge(self.P0, sheet, how='outer', on=[self.col_GenomeStructure, 
					self.col_RepeatRegion, self.col_ORF, self.col_DumaPosition, self.col_DumaSeq])
			sheet = sheet[ sheet[self.col_Sequence] != '-' ]
			self.BPRaw.append(sheet)
			self.BPRawLength.append(len(sheet))
		
		self.Dumas = self.P0[ self.P0[self.col_DumaSeq] != '-' ]
		infolog("mergesequence end")
		
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

	def get_Number_of_GPS(self, BP, s1, s2):
		"""
			s1, s2 범위의 maf값을 갖는 base pair의 수, 길이, minor 값을 
		"""
		from numpy import divide, logical_and
		maf_ = divide(BP[['minor']], BP[['sum']]) * 100
		idx = logical_and(maf_>=s1, maf_<s2)
		idx = idx.values.tolist()
		return maf_[idx].sum()[0], len(BP[idx]), BP[['minor']][idx]

	def save_data(self, book, filename):
		import pickle as pk
		params = {}
		params['filename'] = book.filename
		params['nsheets'] = book.nsheets
		params['constraints'] = book.constraints
		params['sheet_list'] = book.sheet_list
		params['sheets'] = book.sheets
		params['col_DumaPosition'] = book.col_DumaPosition
		params['col_DumaSeq'] = book.col_DumaSeq
		params['col_Sequence'] = book.col_Sequence
		params['col_GenomeStructure'] = book.col_GenomeStructure
		params['col_RepeatRegion'] = book.col_RepeatRegion
		params['col_ORF'] = book.col_ORF
		params['col_Position'] = book.col_Position
		params['col_basepair'] = book.col_basepair
		params['GenomeStructure'] = book.GenomeStructure
		params['RepeatRegion'] = book.RepeatRegion
		params['ORF'] = book.ORF
		params['NCR'] = book.NCR
		params['BPRaw'] = book.BPRaw
		params['BPRawLength'] = book.BPRawLength
		params['BP35'] = book.BP35
		params['BPxMajor'] = book.BPxMajor
		params['BPxMinor'] = book.BPxMinor
		params['P0'] = book.P0
		params['Dumas'] = book.Dumas
		with open(filename, 'wb') as f:
			pk.dump(params, f)

	@staticmethod
	def load_data(filename):
		import pickle as pk
		params = None
		with open(filename, 'rb') as f:
			params = pk.load(f)
		book = Book()
		book.filename = params['filename']
		book.nsheets = params['nsheets']
		book.constraints = params['constraints']
		book.sheet_list = params['sheet_list']
		book.sheets = params['sheets']
		book.col_DumaPosition = params['col_DumaPosition']
		book.col_DumaSeq = params['col_DumaSeq']
		book.col_Sequence = params['col_Sequence']
		book.col_GenomeStructure = params['col_GenomeStructure']
		book.col_RepeatRegion = params['col_RepeatRegion']
		book.col_ORF = params['col_ORF']
		book.col_Position = params['col_Position']
		book.col_basepair = params['col_basepair']
		book.GenomeStructure = params['GenomeStructure']
		book.RepeatRegion = params['RepeatRegion']
		book.ORF = params['ORF']
		book.NCR = params['NCR']
		book.BPRaw = params['BPRaw']
		book.BPRawLength = params['BPRawLength']
		book.BP35 = params['BP35']
		book.BPxMajor = params['BPxMajor']
		book.BPxMinor = params['BPxMinor']
		book.P0 = params['P0']
		book.Dumas = params['Dumas']
		return book