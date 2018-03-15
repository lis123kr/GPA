from openpyxl import Workbook
import numpy as np
import pandas as pd
import logging, time

class Analyzer(object):
	def __init__(self, book):
	"""		
		This program was made by Python, PyQt5, Pyinstaller, openpyxl, pandas, numpy.
		If you want to make excutable files(.exe) using Pyinstaller, you must install above modules.

		# Developed by IL Seop Lee, Software Engineering Department in CBNU, 2017	
		# Refactoring, 2018-03		
	"""
		self.book = book

		self.book.initialize()
		
		# get raw data of full basepair sequence
		self.book.mergesheets()

	def init_full(self, book):
		for bp in book.BPRaw:
			book.BPRawLength.append(len(bp))

			# 염기 A, G, C, T의 합계 컬럼 추가
			bp["sum"] = bp[book.col_basepair].sum(axis=1)

			# 35bp
			# col_basepair 가 리스트라 문제가 생길듯
			bp = bp[[book.col_GenomeStructure, book.col_RepeatRegion, 
					book.col_ORF, book.Sequence, book.col_basepair, "sum"]]
			# sum 값이 35 이상인 데이터 추출
			# constraions = 35.0
			bp = bp[ bp["sum"] >= book.constraints ]
			pmajor, pminor = book.get_major_minor(bp[book.col_basepair])
			# major, minor의 값을 추출
			book.BPxMajor.append(pmajor)
			book.BPxMinor.append(pminor)
			bp['minor'] = pminor
			bp['major'] = pmajor
			
			book.BP35.append(bp)			

		# major, minor의 인덱스 컬럼 추가
		for bp in book.BP35:
			argx = -bp[book.col_basepair] 	# A, G, C, T의 순으로 우선순위를 위해 음수로 정렬
			bps = argx.values.argsort(axis=1)
			bp['minor_idx'] = pd.DataFrame(bps[:,1], index=[bp.index])
			bp['major_idx'] = pd.DataFrame(bps[:,0], index=[bp.index])

	def init_Minor(self, type_, book):
		PP_Raw = []
		PP_Length = []
		PP_sheets = []
		for i, x in enumerate(book.BPRaw):
			for j, y in enumerate(book.BPRaw):
				if i >= j:
					continue
				PP_sheets.append(book.sheet_list[i]+" -> "+book.sheet_list[j])
				bp, pmajor, pminor = self.Extract_difference_of_minor(x,y,type_, book)

				bp["sum"] = bp[book.col_basepair].sum(axis=1)

				PP_Raw.append(bp)
				PP_Length.append(len(bp))

				bp = bp[[book.col_GenomeStructure, book.col_RepeatRegion, 
					book.col_ORF, book.Sequence, book.col_basepair, "sum"]]

				bp = bp[ bp['sum'] >= book.constraints]
				pmajor = pmajor.loc[bp.index]
				pminor = pminor.loc[bp.index]
				bp['minor'] = pminor
				bp['major'] = pmajor
				book.BPxMajor.append(pmajor)
				book.BPxMinor.append(pminor)				
				book.BP35.append(bp)

		book.BPRaw = PP_Raw
		book.BPRawLength = PP_Length
		book.sheet_list = PP_sheets
		book.nsheets = len(book.sheet_list)

		for bp in book.BP35:
			tmp = -bp
			bps = tmp.values.argsort(axis=1)
			bp['minor_idx'] = pd.DataFrame(bps[:,1], index=[bp.index]) if len(bp)!=0 else 0
			bp['major_idx'] = pd.DataFrame(bps[:,0], index=[bp.index]) if len(bp)!=0 else 0

	def Extract_difference_of_minor(self, x, y, type_, book):
		# x, y : PxRaw (removed '-')
		if len(x) > len(y):
			y = pd.DataFrame(y, index = x.index)
		else:
			x = pd.DataFrame(x, index = y.index)
        
		tmp_x = x[ x[book.col_DumaPosition] == y[book.col_DumaPosition]]
		tmp_y = y[ x[book.col_DumaPosition] == y[book.col_DumaPosition]]
		x_bp = -tmp_x[book.col_basepair]
		x_bps = x_bp.values.argsort(axis=1)
		x_minor_idx = pd.DataFrame(x_bps[:,1], index=[x_bp.index])
		x_major_idx = pd.DataFrame(x_bps[:,0], index=[x_bp.index])
    
		y_bp = -tmp_y[book.col_basepair]
		y_bps = y_bp.values.argsort(axis=1)
		y_minor_idx = pd.DataFrame(y_bps[:,1], index=[y_bp.index])
		y_major_idx = pd.DataFrame(y_bps[:,0], index=[y_bp.index])
    
		tmp_x = tmp_x[ (x_major_idx == y_major_idx).values.tolist() ]
		tmp_y = tmp_y[ (x_major_idx == y_major_idx).values.tolist() ]
    
		x_major, x_minor = book.get_major_minor(tmp_x[book.col_basepair])
		y_major, y_minor = book.get_major_minor(tmp_y[book.col_basepair])
		x_minor = pd.DataFrame(x_minor.values, index = tmp_x.index)
		y_minor = pd.DataFrame(y_minor.values, index = tmp_y.index)

		x_sum = pd.DataFrame(tmp_x[book.col_basepair].sum(axis=1), index = tmp_x.index)
		y_sum = pd.DataFrame(tmp_y[book.col_basepair].sum(axis=1), index = tmp_y.index)
    
		x_maf = np.divide(x_minor, x_sum) * 100
		y_maf = np.divide(y_minor, y_sum) * 100
		if type_ == 'ASC':
			return tmp_x[((y_maf - x_maf) >= 5.0).values] , x_major.loc[tmp_x.index], x_minor.loc[tmp_x.index]
		else:
			return tmp_x[((x_maf - y_maf) >= 5.0).values] , x_major.loc[tmp_x.index], x_minor.loc[tmp_x.index]

	def Analyze(self, Analyze_type, book):
		try:
			self.Analyze_type = Analyze_type
			if Analyze_type == "Difference_of_Minor":
				self.s1 = [ 0, 2.5, 5, 15, 25, 5]
				self.s2 = [ 2.5, 5, 15, 25, 51.0, 51.0]
				types_ = ["ASC", "DESC"]
				for i in range(0,2):
					wb = Workbook()
					ws1 = wb.active
					ws2 = wb.create_sheet("GPS")
					ws3 = wb.create_sheet("Genome structure")
					ws4 = wb.create_sheet("ORF")
					ws5 = wb.create_sheet("NCR")
					ws6 = wb.create_sheet("Base composition_1")
					ws7 = wb.create_sheet("Base composition_2")

					self.init_Minor(types_[i], book)
					# self.sheet1_m(ws1, types_[i], book)
					self.sheet2(ws2, book)			
					self.sheet3(ws3, book)			
					self.sheet4_5(ws4, "ORF", book)			
					self.sheet4_5(ws5, "NCR", book)			
					self.sheet6(ws6, book)			
					self.sheet7(ws7, book)
					wb.save("["+types_[i]+"분석]" + book.filename + '.xlsx')
			else: # full
				self.s1 = [ 2.5, 5, 15, 25, 5]
				self.s2 = [ 5, 15, 25, 51.0, 51.0]
				wb = Workbook()
				ws1 = wb.active
				ws2 = wb.create_sheet("GPS")
				ws3 = wb.create_sheet("Genome structure")
				ws4 = wb.create_sheet("ORF")
				ws5 = wb.create_sheet("NCR")
				ws6 = wb.create_sheet("Base composition_1")
				ws7 = wb.create_sheet("Base composition_2")
				self.init_full(book)

				self.sheet1(ws1, book)				
				self.sheet2(ws2, book)			
				self.sheet3(ws3, book)			
				self.sheet4_5(ws4, "ORF", book)			
				self.sheet4_5(ws5, "NCR", book)			
				self.sheet6(ws6, book)			
				self.sheet7(ws7, book)
				wb.save("[분석]" + book.filename + '.xlsx')
			book.P0.to_excel('[통합]'+book.filename+'.xlsx', index=False)
			return "Success"
		except PermissionError:
			return "Permission"
		except Exception as e:
			logging.error("{0} {1}".format(time.time(), e))
			return "Error"

	def get_Number_of_GPS(self, BP, s1, s2):
		"""
			s1, s2 범위의 maf값을 갖는 base pair의 수, 길이, minor 값을 
		"""
		maf_ = np.divide(BP[['minor']], BP[['sum']]) * 100
		idx = np.logical_and(maf_>=s1, maf_<s2)
		idx = idx.values.tolist()
		return maf_[idx].sum()[0], len(BP[idx]), BP[['minor']][idx]

	# def get_level_major_minor(self, major, minor, sum_, s1, s2):
	# 	maf_ = np.divide(minor, sum_) * 100
	# 	idx = np.logical_and(maf_ >= s1, maf_ < s2)
	# 	idx = idx.values.tolist()
	# 	return pd.DataFrame(major[idx].values, index=[sum_[idx].index]), pd.DataFrame(minor[idx].values, index=[sum_[idx].index])
		
	# # genome, minor, sum 병합 함수
	# def merge_Genome_Structure(self, Passage, minor_, sum_, title):
	# 	df_gn = Passage[[title]]
	# 	df_gn = df_gn[ Passage["sum"] >= 35.0 ]
	# 	df_gn_minor = pd.concat([pd.DataFrame(df_gn.values), pd.DataFrame(minor_.values), pd.DataFrame(sum_.values)], axis=1)
	# 	return pd.DataFrame(df_gn_minor.values, columns=[title, "minor", "sum"], index=[sum_.index])
		
	def insert_value_in_cell(self,ws, rows, col, BP, minor_, maf, colname, types):
		"""
			한 column에 데이터 삽입
			- rows : 데이터를 삽입할 row list
			- BP : 'sum' >= 35.0 처리된 DataFrame
			- minor_ : s1~s2
			- maf : s / l
			- types : GPS or MAF
		"""
		index = book.GenomeStructure if colname == book.col_GenomeStructure else book.RepeatRegion
	
		# For Minor Exception
		if len(BP) == 0:
			for i in range(0, len(rows)):
				ws[col+str(rows[i])] = 0 if types == "GPS" else 'N/A'
			ws[col+str(rows[len(rows)-1] + 1)] = 0 if types == "GPS" else 'N/A'
		else:
			# merged = self.merge_Genome_Structure(Passage, minor_, sum_, colname)
			# _, tm_minor = self.get_level_major_minor(minor_, minor_, sum_, s1,s2)
			mid_rows = BP.loc[ minor_.index ]
			cnt_sum=0;
			for i in range(0, len(rows)): # rows[i], index[i]		
				# 각 index 별로 그룹핑
				mrows = mid_rows [ mid_rows[ colname ] == index[i] ]
				if types == "GPS":
					ws[col+str(rows[i])] = len(mrows)
					cnt_sum+=len(mrows)
				elif types == "MAF":
					if len(mrows) != 0:
						maf_ = np.divide(mrows[["minor"]], mrows[["sum"]]) * 100
						ws[col+str(rows[i])] = str(round(maf_.sum()[0] / len(maf_),3)) + '%'
					else:
						ws[col+str(rows[i])] = 'N/A'
			if types == "GPS":
				ws[col+str(rows[len(rows)-1] + 1)] = cnt_sum
			elif types == "MAF":
				# s, l = self.get_Number_of_GPS(BP, s1, s2)
				ws[col+str(rows[len(rows)-1] + 1)] = str(round(maf, 3)) + '%' if l is not 0 else 'N/A'

	def next_col(self, asc):
		asc2 = list(asc)
		if len(asc2) == 1 and asc2[0] != 'Z':
			return chr(ord(asc2[0])+1)
		elif len(asc2) == 1 and asc2[0] == 'Z':
			return 'AA'
		else:
			if asc[-1] != 'Z':
				return asc[:-1] + chr(ord(asc[-1])+1)
			else:
				p = None
				for i in range(-1, -len(asc)-1, -1):
					p = i
					if asc2[i] == 'Z':
						asc2[i] = 'A'
					else:
 						break
				if p == -len(asc) and asc[p] == 'Z':
					return 'A'+''.join(asc2)
				else:
					asc = ''.join(asc2)
					return asc[:p] + chr(ord(asc[p])+1) + asc[p+1:]

	# def sheet1_m(self, ws, types_, book):
	# 	logging.info("{0} Start Sheet1_m".format(time.time()))
	# 	ws.title = "Polymorphic site"
	# 	# A col
	# 	ws['A1'] = types_
	# 	ws.merge_cells(start_row=1, start_column=1, end_row=2, end_column=2)
		
	# 	ws['A3'] = self.filename
	# 	ws.merge_cells(start_row=3, start_column=1, end_row=2+self.length, end_column=1)
	# 	# C col
	# 	ws['C1'] = "Genome length (bp)"
	# 	ws.merge_cells(start_row=1, start_column=3, end_row=2, end_column=3)
	# 	# D col
	# 	ws['D1'] = "Genome length (35이상)"
	# 	ws.merge_cells(start_row=1, start_column=4, end_row=2, end_column=4)
		
	# 	ws['E1'] = "Average of MAF"
	# 	ws.merge_cells(start_row=1, start_column=5, end_row=2, end_column=5)

	# 	ws['F1'] = "Number of polymorphic sites"
	# 	ws.merge_cells(start_row=1, start_column=6, end_row=1, end_column=11)

	# 	ws['F2'] = "0≤n<2.5"
	# 	ws['G2'] = "2.5≤n<5"
	# 	ws['H2'] = "5≤n<15"
	# 	ws['I2'] = "15≤n<25"
	# 	ws['J2'] = "25≤n"
	# 	ws['K2'] = "5≤n"
	# 	for i in range(self.length):
	# 		ws['B' + str(i+3)] = self.sheet_names[i]
	# 		ws['C' + str(i+3)] = self.PxLength[i]
	# 		ws['D' + str(i+3)] = len(self.P_Raw35[i])
	# 		ws['E' + str(i+3)] = str(round(self.PxMAF[i], 3)) + '%'
	# 		ws['F' + str(i+3)] = self.get_Number_of_GPS(self.PxMinor[i], self.PxSum[i], 0.0, 2.5)[1] if len(self.PxMinor[i])!=0 else 0
	# 		ws['G' + str(i+3)] = self.get_Number_of_GPS(self.PxMinor[i], self.PxSum[i], 2.5, 5.0)[1] if len(self.PxMinor[i])!=0 else 0
	# 		ws["H" + str(i+3)] = self.get_Number_of_GPS(self.PxMinor[i], self.PxSum[i], 5.0, 15.0)[1] if len(self.PxMinor[i])!=0 else 0
	# 		ws["I" + str(i+3)] = self.get_Number_of_GPS(self.PxMinor[i], self.PxSum[i], 15.0, 25.0)[1] if len(self.PxMinor[i])!=0 else 0
	# 		ws["J" + str(i+3)] = self.get_Number_of_GPS(self.PxMinor[i], self.PxSum[i], 25.0, 51.0)[1] if len(self.PxMinor[i])!=0 else 0
	# 		ws["K" + str(i+3)] = self.get_Number_of_GPS(self.PxMinor[i], self.PxSum[i], 5.0, 51.0)[1] if len(self.PxMinor[i])!=0 else 0
	# 	logging.info("{0} End Sheet1_m".format(time.time()))

	def sheet1(self,ws, book):
		logging.info("{0} Start Sheet1".format(time.time()))
		ws.title = "Polymorphic site"
		ws['A1'] = "strain"
		ws['D1'] = "Average of MAF"
		ws['E1'] = "Number of polymorphic sites"
		ws.merge_cells(start_row=1, start_column=5, end_row=1, end_column=10)
		ws['E2'] = "2.5≤n<5"
		ws['F2'] = "5≤n<15"
		ws['G2'] = "15≤n<25"
		ws['H2'] = "25≤n"
		ws['I2'] = "5≤n"
		# A col
		ws.merge_cells(start_row=1, start_column=1, end_row=2, end_column=2)
		ws['C1'] = "Genome length (bp)"
		ws.merge_cells(start_row=1, start_column=3, end_row=2, end_column=3)		
		# D col
		ws.merge_cells(start_row=1, start_column=4, end_row=2, end_column=4)
		ws['A3'] = book.filename
		ws.merge_cells(start_row=3, start_column=1, end_row=2+book.nsheets, end_column=1)
		for i in range(book.nsheets):
			logging.info("{0} Writing {1} sheet".format(time.time(), book.sheet_list[i]))
			ws['B' + str(i+3)] = book.sheet_list[i]
			ws['C' + str(i+3)] = book.BPRawLength[i]
			s, l, _ = self.get_Number_of_GPS(BP35[i], 5.0, 51.0)
			ws['D' + str(i+3)] = str(round(s / l, 3)) + '%'
			ws['E' + str(i+3)] = self.get_Number_of_GPS(BP35[i], 2.5, 5.0)[1]
			ws["F" + str(i+3)] = self.get_Number_of_GPS(BP35[i], 5.0, 15.0)[1]
			ws["G" + str(i+3)] = self.get_Number_of_GPS(BP35[i], 15.0, 25.0)[1]
			ws["H" + str(i+3)] = self.get_Number_of_GPS(BP35[i], 25.0, 51.0)[1]
			ws["I" + str(i+3)] = self.get_Number_of_GPS(BP35[i], 5.0, 51.0)[1]
		logging.info("{0} End Sheet1".format(time.time()))

	def sheet2(self,ws, book):
		logging.info("{0} Start Sheet2".format(time.time()))
		ws['A1'] = "strain"
		ws.merge_cells(start_row=1, start_column=1, end_row=2, end_column=2)
		ws['C1'] = "Genome length (bp)"
		ws.merge_cells(start_row=1, start_column=3, end_row=2, end_column=3)
		ws['D1'] = "Average of MAF at GPS"
		ws.merge_cells(start_row=1, start_column=4, end_row=1, end_column=9)
		ws['D2'] = "2.5≤n<5"
		ws['E2'] = "5≤n<15"
		ws['F2'] = "15≤n<25"
		ws['G2'] = "25≤n"
		ws['H2'] = "5≤n"
		ws['A3'] = book.filename
		ws.merge_cells(start_row=3, start_column=1, end_row=2+book.nsheets, end_column=1)

		for i in range(book.nsheets):
			logging.info("{0} Writing {1} sheet".format(time.time(), book.sheet_list[i]))

			ws['B' + str(i+3)] = book.sheet_list[i]
			ws['C' + str(i+3)] = book.BPRawLength[i]
			
			col = 'D'
			for r1, r2 in zip(self.s1, self.s2):
				s, l, _ = self.get_Number_of_GPS(BP35[i], r1, r2) if len(book.BPxMinor[i])!=0 else (0,0)
				ws[col + str(i+3)] = str(round(s/l, 3))+'%' if l is not 0 else 'N/A'
				col = self.next_col(col)

			if self.Analyze_type == "Difference_of_Minor":
				ws['I2'] = "0≤n<2.5"
				s, l, _ = self.get_Number_of_GPS(BP35[i], 0, 2.5) if len(book.BPxMinor[i])!=0 else (0,0)
				ws['I' + str(i+3)] = str(round(s/l,3)) + '%' if l is not 0 else 'N/A'

		logging.info("{0} End Sheet2".format(time.time()))

	def Full_Seq_in_sheet3(self, ws, col, book):
		logging.info("{0} Start Full_Seq_in_sheet3".format(time.time()))
		# Writing list of GenomeStructure & RepeatRegion, NCR & ORF
		for r in range(4, 4+len(book.GenomeStructure)):
			ws[col+str(r)] = book.GenomeStructure[r-4]
		ws[col+str(4+len(book.GenomeStructure))] = 'Total'
		for r in range(0, len(book.RepeatRegion)):
			row = r+5+len(book.GenomeStructure)
			ws[col+str(row)] = book.RepeatRegion[r]
		ws[col+str(5+len(book.GenomeStructure)+len(book.RepeatRegion))] = 'Total'

		rs_ = 5+len(book.GenomeStructure)+len(book.RepeatRegion)
		ws[col+str(rs_+1)] = 'NCR'
		ws[col+str(rs_+2)] = 'ORF'

		# Writing the count of each sheets
		col = self.next_col(col)
		for i in range(0, book.nsheets):
			ws[col+str(2)] = book.sheet_list[i]

			# 35 이하 포함
			ws[col+str(3)] = '(BP_full)Length'
			sum_ = 0
			for r in range(4, 4+len(book.GenomeStructure)):
				tmp = len(book.BPRaw[i][ book.BPRaw[i][ book.col_GenomeStructure] == book.GenomeStructure[r-4]]) if len(book.BPRaw[i])!=0 else 0
				ws[col+str(r)] = tmp
				sum_ += tmp
			ws[col+str(4+len(book.GenomeStructure))] = sum_
			sum_=0
			for r in range(0, len(book.RepeatRegion)):
				row = r+5+len(book.GenomeStructure)
				tmp = len(book.BPRaw[i][ book.BPRaw[i][ book.col_RepeatRegion] == book.RepeatRegion[r]]) if len(book.BPRaw[i])!=0 else 0
				ws[col+str(row)] = tmp
				sum_ += tmp
			ws[col+str(5+len(book.GenomeStructure)+len(book.RepeatRegion))] = sum_

			# ORF & NCR
			ws[col+str(rs_+1)] = len(book.BPRaw[i][ book.BPRaw[i][ book.col_ORF].isin(book.ORF)]) if len(book.BPRaw[i])!=0 else 0
			ws[col+str(rs_+2)] = len(book.BPRaw[i][ book.BPRaw[i][ book.col_ORF].isin(book.NCR)]) if len(book.BPRaw[i])!=0 else 0

			# sum of basepair 35 이상
			col = self.next_col(col)
			ws[col+str(3)] = '(BP_35이상)Length'
			sum_ = 0
			for r in range(4, 4+len(book.GenomeStructure)):
				tmp = len(book.BP35[i][ book.BP35[i][ book.col_GenomeStructure] == book.GenomeStructure[r-4]]) if len(book.BPRaw[i])!=0 else 0
				ws[col+str(r)] = tmp
				sum_ += tmp
			ws[col+str(4+len(book.GenomeStructure))] = sum_
			sum_=0
			for r in range(0, len(book.RepeatRegion)):
				row = r+5+len(book.GenomeStructure)
				tmp = len(book.BP35[i][ book.BP35[i][ book.col_RepeatRegion] == book.RepeatRegion[r]]) if len(book.BP35[i])!=0 else 0
				ws[col+str(row)] = tmp
				sum_ += tmp
			ws[col+str(5+len(book.GenomeStructure)+len(book.RepeatRegion))] = sum_

			ws[col+str(rs_+1)] = len(book.BP35[i][ book.BP35[i][ book.col_ORF ].isin(book.ORF)]) if len(book.BP35[i])!=0 else 0
			ws[col+str(rs_+2)] = len(book.BP35[i][ book.BP35[i][ book.col_ORF ].isin(book.NCR)]) if len(book.BP35[i])!=0 else 0
			col = self.next_col(col)
		logging.info("{0} End Full_Seq_in_sheet3".format(time.time()))

	def sheet3(self,ws, book):
		logging.info("{0} Start sheet3".format(time.time()))
		for i in range(0, len(self.s1)):
			logging.info("{0} Writing {1} ~ {2}".format(time.time(), self.s1[i], self.s2[i]))

			# r : (s1,s2)의 범위 별 데이터의 row 변수
			r = i*(15+len(book.GenomeStructure) + len(book.RepeatRegion))+ 1
			ws['A'+str(r)] = "Region"
			ws.merge_cells(start_row=r, start_column=1, end_row=r+2, end_column=2)
			ws['C'+str(r)] = "Length (bp)"
			ws.merge_cells(start_row=r, start_column=3, end_row=r+2, end_column=3)

			ws['A'+str(r+3)] = "Genome Structure"
			ws.merge_cells(start_row=r+3, start_column=1, end_row=r+3+len(book.GenomeStructure), end_column=1)
			ws['A'+str(r+10)] = "Repeat region"
			ws.merge_cells(start_row=r+4+len(book.GenomeStructure), start_column=1, end_row=r+4+len(book.GenomeStructure)+len(book.RepeatRegion), end_column=1)
			
			# B, C 컬럼 데이터 : GenomeStructure, RepeatRegion 종류와 길이
			# G_rows : GenomeStructure데이터를 넣기 위한 row 범위
			# R_rows : RepeatRegion데이터를 넣기 위한 row 범위
			G_rows = list(range(r+3, r+3+len(book.GenomeStructure)))
			cnt_num = 0
			for ix in range(0, len(book.GenomeStructure)):
				ws['B' + str(G_rows[ix])] = book.GenomeStructure[ix]
				nlen = len(book.P0[ book.P0[ book.col_GenomeStructure ] == book.GenomeStructure[ix]])
				ws['C' + str(G_rows[ix])] = nlen
				cnt_num += nlen
			ws['B' + str(r+3+len(book.GenomeStructure))] = "Total"
			ws['C' + str(r+3+len(book.GenomeStructure))] = cnt_num
			cnt_num = 0

			nr = r+4+len(book.GenomeStructure)+len(book.RepeatRegion)
			R_rows = list(range(r+4+len(book.GenomeStructure), nr))
			for ix in range(0, len(book.Region_idx)):
				ws['B' + str(R_rows[ix])] = book.RepeatRegion[ix]
				nlen = len(book.P0[ book.P0[ book.col_RepeatRegion ] == book.RepeatRegion[ix]])
				ws['C'+str(R_rows[ix])] = nlen
				cnt_num += nlen
			ws['B'+str(nr)] = 'Total'
			ws['C' + str(nr)] = cnt_num

			## ORF NCR
			ws["A"+str(nr+1)] = "ORF"
			ws["A"+str(nr+2)] = "NCR"
			# ORF NCR Full Length
			ws['C'+str(nr+1)] = len(book.P0[ book.P0[book.col_ORF].isin(book.ORF)])
			ws['C'+str(nr+2)] = len(book.P0[ book.P0[book.col_ORF].isin(book.NCR)])
			# end B, C columns
			
			# ncol = 'D'
			# ncol : Column of 'Number of GPS' is need when calculating 'Number of GPS / Length'
			ncol = col = 'D'

			for c in range(0, book.nsheets):
				ws[col+str(r+2)] = book.sheet_list[c]
				# 35 이상 
				# px_orf_merged = self.merge_Genome_Structure(self.P_Raw[c], self.PxMinor[c],
				# 				self.PxSum[c], self.Dorf)
				# _, px_minor = self.get_level_major_minor(self.PxMinor[c], self.PxMinor[c],
				# 				self.PxSum[c], self.s1[i], self.s2[i])
				s, l, minor_ = self.get_Number_of_GPS(book.BP35[c], self.s1[i], self.s2[i])

				tx_rows = BP35[c].loc[ minor_.index ]
				# tx_rows = px_orf_merged.loc[ px_minor.index ]
				tx_orf = tx_rows[ tx_rows[ book.col_ORF ].isin(book.ORF)]
				tx_ncr = tx_rows[ tx_rows[ book.col_ORF ].isin(book.NCR)]

				ws[col+str(r)] = "Number of GPS"
				ws.merge_cells(start_row=r, start_column=4, end_row=r, end_column=4 + book.nsheets - 1)
				ws[col+str(r+1)] = book.filename
				ws.merge_cells(start_row=r+1, start_column=4, end_row=r+1, end_column=4 + book.nsheets - 1)

				self.insert_value_in_cell(ws, rows, col, 
					book.BP35[c], minor_, s/l,  book.col_GenomeStructure , "GPS")
				self.insert_value_in_cell(ws, nrows, col, 
					book.BP35[c], minor_, s/l,  book.col_GenomeStructure , "GPS")
				ws[col+str(nr+1)] = len(tx_orf)
				ws[col+str(nr+2)] = len(tx_ncr)

				col = self.next_col(col)

			# "Average MAF at GPS"
			for c in range(0, book.nsheets):
				ws[col+str(r)] = "Average MAF at GPS"
				ws.merge_cells(start_row=r, start_column=4 + book.nsheets, end_row=r, end_column=4 + 2 * book.nsheets - 1)
				ws[col+str(r+1)] = book.filename
				ws.merge_cells(start_row=r+1, start_column=4 + book.nsheets, end_row=r+1, end_column=4 + 2 * book.nsheets - 1)

				self.insert_value_in_cell(ws, rows, col, book.BP35[c], minor_, s/l,  book.col_RepeatRegion , "MAF")
				self.insert_value_in_cell(ws, nrows, col, book.BP35[c], minor_, s/l,  book.col_RepeatRegion , "MAF")

				ws[col+str(nr+1)] = str(round((np.divide(tx_orf[["minor"]], book.BP35[c].loc[ tx_orf.index]) * 100).sum()[0]/len(tx_orf), 3)) +'%' if len(tx_orf) is not 0 else 'N/A'
				ws[col+str(nr+2)] = str(round((np.divide(tx_ncr[["minor"]], book.BP35[c].loc[ tx_ncr.index]) * 100).sum()[0]/len(tx_ncr), 3)) +'%' if len(tx_ncr) is not 0 else 'N/A'
				
				col = self.next_col(col)

			# "Number of GPS / length"
			for c in range(0, book.nsheets):
				####
				ws[col+str(r)] = "Number of GPS / length"
				ws.merge_cells(start_row=r, start_column=4 + 2 *book.nsheets, end_row=r, end_column=4 + 3 *book.nsheets - 1)
				ws[col+str(r+1)] = book.filename
				ws.merge_cells(start_row=r+1, start_column=4 + 2 *book.nsheets, end_row=r+1, end_column=4 + 3 *book.nsheets - 1)
				ridx = sum_ = 0
				
				for rs in range(r+3, r+5+len(book.GenomeStructure)+len(book.RepeatRegion)):
					if ridx < len(book.GenomeStructure):
						llen_ = len(book.BP35[c][ book.BP35[c][ book.col_GenomeStructure] == book.GenomeStructure[ridx]])
						sum_ += llen_
					elif ridx == len(book.GenomeStructure):
						llen_ = sum_
						sum_ = 0
					elif len(book.GenomeStructure) < ridx and ridx < len(book.GenomeStructure)+1+len(book.RepeatRegion):
						llen_ = len(book.BP35[c][ book.BP35[c][ book.col_RepeatRegion] == book.RepeatRegion[ridx-1-len(book.GenomeStructure)]])
						sum_ += llen_
					else:
						llen_ = sum_

					# ncol = chr(ord(col)-self.length*2)
					ws[col + str(rs)] = ws[ncol+str(rs)].value / llen_ if llen_ is not 0 else 'N/A'					
					ridx = ridx + 1

				# ORF & NCR
				for rs in range(2):
					rs_ = r+5+len(book.GenomeStructure)+len(book.RepeatRegion) + rs
					if rs == 0:
						tmp = len(book.BP35[c][ book.BP35[c][book.col_ORF].isin(book.ORF)])
						ws[col+str(rs_)] = ws[ncol+str(rs_)].value / tmp if tmp is not 0 else 'N/A'
							
					else:
						tmp = len(book.BP35[c][ book.BP35[c][book.col_ORF].isin(book.NCR)])
						ws[col+str(rs_)] = ws[ncol+str(rs_)].value / tmp if tmp is not 0 else 'N/A'

				ncol = self.next_col(ncol)
				col = self.next_col(col)				

			col = self.next_col(col)
			ws[col+str(r+3)] = str(self.s1[i]) + "~" + str(self.s2[i]) + "%" if self.s2[i] is not 51.0 else str(self.s1[i]) + "% 이상"			
		logging.info("{0} finished writing GenomeStructure & RepeatRegion".format(time.time()))

		#### Full Sequence				
		col = self.next_col(self.next_col(col))		
		self.Full_Seq_in_sheet3(ws, col, book)
		logging.info("{0} End sheet3".format(time.time()))

	def sheet4_5(self, ws, title, book):
		logging.info("{0} Start sheet {1}".format(time.time(), 4 if title is "ORF" else 5))
		ws['B2'] = book.filename
		ws["C2"] = "(BP_full)Length"
		ws.merge_cells(start_row=2, start_column=3, end_row=2, end_column=3+book.nsheets-1)
		ws["B3"] = title
		col = 'C'
		for i in range(0, book.nsheets):
			ws[col + str(3)] = book.sheet_list[i]
			col = self.next_col(col)		

		# Full - 35이하 포함
		col = 'B'
		orf_ = book.ORF if title is 'ORF' else book.NCR
		rows = list(range(4, len(orf_) + 4))
		for c in range(0, book.nsheets+1):
			cnt_ = 0
			for i in range(0, len(orf_)):
				if c is 0:
					ws[col+str(rows[i])] = orf_[i]
				else:
					tmp = len(book.BPRaw[c-1][ book.BPRaw[c-1][book.col_ORF] == orf_[i] ])
					ws[col+str(rows[i])] = tmp
					cnt_ += tmp
			ws[col + str(4+len(orf_))] = cnt_ if c is not 0 else 'total'
			col = self.next_col(col)

		# s1~s2 범위의 ORF & NCR 계산
		# px_orf_merged = list()
		# for n in range(0, book.nsheets):
		# 	px_orf_merged.append(self.merge_Genome_Structure(self.P_Raw[n], self.PxMinor[n], self.PxSum[n], self.Dorf))

		col = chr(ord('A')+3+book.nsheets)
		for i in range(0, len(self.s1)):
			co = (2+3*book.nsheets) * i + (5+book.nsheets)
			ws[col+str(2)] = str(self.s1[i]) + '~' + str(self.s2[i]) +'%' if self.s2[i] is not 51.0 else str(self.s1[i]) + '%~'
			ws[col+str(3)] = title # NCR or ORF
			for ix in range(0, len(orf_)):
				ws[col+str(rows[ix])] = orf_[ix]
			ws[col+str(len(orf_) + 4)] = 'total'

			# Number of GPS of each sheets
			gps = list()
			col = self.next_col(col)			
			ws[col+str(2)] = "Number of GPS"
			ws.merge_cells(start_row=2, start_column=co, end_row=2, end_column=co+book.nsheets-1)
			for n in range(0, book.nsheets):
				ws[col+str(3)] = book.sheet_list[n]
				_, _, tx_minor = self.get_Number_of_GPS(book.BP35[n], self.s1[i], self.s2[i])
				tx_rows = BP35[n].loc[ tx_minor.index ]

				g = []
				cnt_ = 0
				for ix in range(0, len(orf_)):
					mrows = tx_rows[ tx_rows[book.col_ORF] == orf_[ix] ]
					ws[col+str(rows[ix])] = len(mrows)
					cnt_ += len(mrows)
					g.append(len(mrows))
				ws[col+str(len(orf_) + 4)] = cnt_
				gps.append(g)
				col = self.next_col(col)

			# Average MAF at GPS  of each sheets
			ws[col+str(2)] = "Average MAF at GPS"
			ws.merge_cells(start_row=2, start_column=co+book.nsheets, end_row=2, end_column=co+2*book.nsheets-1)
			for n in range(0, book.nsheets):
				ws[col+str(3)] = book.sheet_list[n]
				_, _, tx_minor = self.get_Number_of_GPS(book.BP35[n], self.s1[i], self.s2[i])
				tx_rows = BP35[n].loc[ tx_minor.index ]
				for ix in range(0, len(orf_)):
					mrows = tx_rows[ tx_rows[book.col_ORF] == orf_[ix] ]
					if len(mrows) is 0:
						ws[col+str(rows[ix])] = 'N/A'
					else:
						maf_ = np.divide(mrows[["minor"]], mrows[["sum"]]) * 100
						val = maf_.sum()[0] / len(maf_)
						ws[col+str(rows[ix])] = str(round(val, 3)) + '%'
				col = self.next_col(col)

			# "Number of GPS / Length" of each sheets
			ws[col+str(2)] = "Number of GPS / Length"
			ws.merge_cells(start_row=2, start_column=co+2*book.nsheets, end_row=2, end_column=co+3*book.nsheets-1)
			ncol = 'C'
			for n in range(0, book.nsheets):
				ws[col+str(3)] = book.sheet_list[n]
				for ix in range(0, len(orf_)):
					l_ = int(str(ws[ncol+str(rows[ix])].value))
					ws[col+str(rows[ix])] = gps[n][ix] / l_ if l_ is not 0 else 'N/A'
				ncol = self.next_col(ncol)
				col = self.next_col(col)

			# end for s1
			col = self.next_col(col)
		logging.info("{0} End sheet {1}".format(time.time(), 4 if title is "ORF" else 5))

	def sheet6(self,ws, book):
		logging.info("{0} Start sheet6".format(time.time()))
		col = 'B'
		for i in range(0, len(self.s1)):
			ws[col+'2'] = str(self.s1[i]) + '~' + str(self.s2[i]) +'%' if self.s2[i] is not 51.0 else str(self.s1[i]) + '%~'

			# get basepair between s1 and s2
			mxr = []
			for n in range(0, book.nsheets):
				_, _, minor_ = self.get_Number_of_GPS(book.BP35[n], self.s1[i], self.s2[i])
				mxr.append(book.BP35[n].loc[minor_.index])

			for x in range(-1, book.nsheets):
				ws[col+'3'] = book.sheet_list[x] if x is not -1 else "Major/Minor"
				rows = 4
				for a in range(0, len(book.col_basepair)):
					for b in range(0, len(book.col_basepair)):
						if a==b:
							continue
						ws[col+str(rows)] = len(mxr[x][np.logical_and( mxr[x]['major_idx']==a, 
							mxr[x]['minor_idx']==b )]) if x is not -1 else book.col_basepair[a] + '/' + book.col_basepair[b]
						rows = rows + 1
				col = self.next_col(col)
			col = self.next_col(col)
		logging.info("{0} End sheet6".format(time.time()))

	def sheet7(self, ws, book):
		logging.info("{0} Start sheet7".format(time.time()))
		for i in range(0, len(self.s1)):
			r = i * (book.nsheets + 5) + 2
			ws["A"+str(r)] = str(self.s1[i]) + "~" + str(self.s2[i]) + "%" if self.s2[i] is not 51.0 else str(self.s1[i]) + "%~"
			
			ws["B"+str(r)] = "Virus"
			ws.merge_cells(start_row=r, start_column=2, end_row=r+2, end_column=2)

			ws["C"+str(r)] = "Number of GPS"
			ws.merge_cells(start_row=r, start_column=3, end_row=r+2, end_column=3)

			ws["D"+str(r)] = "GPS Mean"
			ws.merge_cells(start_row=r, start_column=4, end_row=r+2, end_column=4)

			ws["E"+str(r)] = "Major"
			ws.merge_cells(start_row=r, start_column=5, end_row=r+1, end_column=5)
			ws["E"+str(r+2)] = "Minor"

			ws['F'+str(r)] = 'A'
			ws.merge_cells(start_row=r, start_column=6, end_row=r+1, end_column=8)
			ws['I'+str(r)] = 'G'
			ws.merge_cells(start_row=r, start_column=9, end_row=r+1, end_column=11)
			ws['L'+str(r)] = 'C'
			ws.merge_cells(start_row=r, start_column=12, end_row=r+1, end_column=14)
			ws['O'+str(r)] = 'T'
			ws.merge_cells(start_row=r, start_column=15, end_row=r+1, end_column=17)

			for x in range(0, book.nsheets):
				row = r+x+3
				s, l, minor_ = self.get_Number_of_GPS(book.BP35[x], self.s1[i], self.s2[i])
				mxr = book.BP35[x].loc[minor_.index]
				# s, l = self.get_Number_of_GPS(self.PxMinor[x], self.PxSum[x], self.s1[i], self.s2[i]) if len(self.PxMinor[x])!=0 else (0,0)
				ws["B"+str(row)] = book.sheet_list[x]
				ws['C'+str(row)] = len(minor_)
				ws['D'+str(row)] = str(round(s/l, 3)) + '%' if l is not 0 else 'N/A'
					
				col = 'F'
				for a in range(0, 4):
					for b in range(0, 4):
						if a == b:
							continue
						else:
							ws[col+str(r+2)] = book.col_basepair[b].lower()
							ws[col+str(row)] = len(mxr[np.logical_and( mxr['major_idx']==a, mxr['minor_idx']==b )])
						col = self.next_col(col)
		logging.info("{0} End sheet7".format(time.time()))
