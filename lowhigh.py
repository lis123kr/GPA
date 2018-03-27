from analyzer import get_Number_of_GPS, next_col, infolog
class LowHigh:
	"""
		analysis of increase or decrease sequence
	"""
	def __init__(self, book):
		infolog("lowhigh init start")
		sheet_names = []
		self.BPmerged = []
		self.BPinc = []
		self.BPdec = []
		BPmaf = []

		for bp in book.BP35:
			#get maf 5%
			s, l, minor_ = get_Number_of_GPS(bp, 5.0, 51.0)
			BPmaf.append(bp.loc[minor_.index])

		from pandas import merge, DataFrame
		from numpy import divide
		for i, x in enumerate(BPmaf):
			for j, y in enumerate(BPmaf):
				continue if i >= j
				sheet_names.append(book.sheet_list[i]+'->'+book.sheet_list[j])
				# low의 maf5%이상과 high의 maf5%이상 값을 병합
				# 둘 중 하나라도 sum >= 35.0 and maf 5% 이상인 값에서 분석을 진행
				merged = merge(x, y, how='outer', on=[book.col_GenomeStructure, book.col_RepeatRegion, 
					book.col_ORF, book.col_DumaPosition, book.col_DumaSeq], right_index=True, left_index=True)

				self.BPmerged.append(merged)

				# dumas position 기준으로 양쪽 위치 추출
				x1 = book.BP35[i][ book.BP35[i][book.col_DumaPosition].isin(merged[book.col_DumaPosition].values.tolist())] if len(merged) is not 0 else DataFrame()
				y1 = book.BP35[j][ book.BP35[j][book.col_DumaPosition].isin(merged[book.col_DumaPosition].values.tolist())]	if len(merged) is not 0 else DataFrame()

				if(len(x1) is 0 or len(y1) is 0):
					self.BPinc.append(DataFrame())
					self.BPdec.append(DataFrame())
					continue

				# 0일수도 예외처리..
				x_maf = divide(x1[['minor']], x1[['sum']]) * 100 
				y_maf = divide(y1[['minor']], y1[['sum']]) * 100

				# 증가하는 위치와 감소하는 위치
				incidx = (y_maf - x_maf >= 5.0).values
				decidx = (x_maf - y_maf >= 5.0).values

				if(len(x_maf) > len(y_maf)):
					self.BPinc.append(x1[incidx])
					self.BPdec.append(x1[decidx])
				else:
					self.BPinc.append(y1[incidx])
					self.BPdec.append(y1[decidx])

		book.sheet_list = sheet_names
		book.nsheets = len(sheet_names)
		infolog("lowhigh init end")

	def PolymorphicSite(self, ws, types_, book):
		infolog("lowhigh sheet1 start")
		ws.title = "INC-Polymorphic site"
		# A col
		ws['A1'] = types_
		ws.merge_cells(start_row=1, start_column=1, end_row=2, end_column=2)
		
		# C col
		ws['C1'] = "Increase in genetic polymrphism (%)"
		ws.merge_cells(start_row=1, start_column=3, end_row=1, end_column=6)

		ws['C2'] = "5≤n<15"
		ws['D2'] = "15≤n<25"
		ws['E2'] = "25≤n"
		ws['F2'] = "Sum"
		from numpy import logical_and
		bp = self.BPinc if types == 'INC' else self.BPdec
		for i in range(book.nsheets):
			ws['A' + str(i+3)] = book.sheet_list[i]
			ws.merge_cells(start_row=i+3, start_column=1, end_row=i+3, end_column=2)

			ws['C' + str(i+3)] = len(bp[i][ logical_and(bp[i] >= 5.0, bp[i] < 15).values]) if len(bp) is not 0 else 0
			ws['D' + str(i+3)] = len(bp[i][ logical_and(bp[i] >= 15.0, bp[i] < 25).values]) if len(bp) is not 0 else 0
			ws['E' + str(i+3)] = len(bp[i][ logical_and(bp[i] >= 25.0).values]) if len(bp) is not 0 else 0
			ws['F' + str(i+3)] = len(bp[i])
		infolog("lowhigh sheet1 end")

	def GenomeStr(self, ws, types_, book):
		ws['A1'] = 'Region'
		ws.merge_cells(start_row=1, start_column=1, end_row=2, end_column=1)
		ws['B1'] = 'Dumas Length'
		ws.merge_cells(start_row=1, start_column=2, end_row=2, end_column=2)

		for r in range(3, 3+len(book.GenomeStructure)):
			ws['A' + str(r)] = book.GenomeStructure[r-3]
			ws['B' + str(r)] = book.Dumas[ book.Dumas[book.col_GenomeStructure] == book.GenomeStructure[r-3] ]

		ws['A' + str(3+len(book.GenomeStructure))] = 'Total'
		ws['B' + str(3+len(book.GenomeStructure))] = len(book.Dumas)

		sum_ = 0
		for r in range(len(book.RepeatRegion)):
			row = r+4+len(book.GenomeStructure)
			ws['A' + str(row)] = book.RepeatRegion[r]
			cnt = book.Dumas[ book.Dumas[book.col_RepeatRegion] == book.RepeatRegion[r] ]
			ws['B' + str(row)] = cnt
			sum_ += cnt

		leng = len(book.GenomeStructure) + len(book.RepeatRegion)
		ws['A' + str(leng+4)] = 'Total'
		ws['B' + str(leng+4)] = sum_
		ws['A' + str(leng+5)] = 'ORF'
		ws['B' + str(leng+5)] = len(book.Dumas[ book.Dumas[book.col_ORF].isin(book.ORF)])
		ws['A' + str(leng+6)] = 'NCR'
		ws['B' + str(leng+6)] = len(book.Dumas[ book.Dumas[book.col_ORF].isin(book.NCR)])

