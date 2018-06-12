from analyzer import next_n_col, next_col, infolog, errorlog# book.get_Number_of_GPS, next_col, infolog
import time
from datetime import datetime
from numpy import divide, logical_and, logical_or

class FullSeq:
	"""
		analysis of full sequence
	"""
	def __init__(self, book):	
		self.s1 = [ 2.5, 5, 15, 25, 5]
		self.s2 = [ 5, 15, 25, 51.0, 51.0]	

	def maf5list(self, ws, book):
		from analyzer import dataframe_to_rows
		from pandas import merge
		infolog("writing list of data")
		P0, merged = None, None
		dumascol = [book.col_GenomeStructure, book.col_RepeatRegion, book.col_ORF, book.col_DumaPosition, book.col_DumaSeq]
		for i in range(book.nsheets):
			book.BP35[i]['maf'] = divide(book.BP35[i][['minor']], book.BP35[i][['sum']]) * 100
			if merged is not None:
				merged = merge(merged, book.BP35[i][ book.BP35[i]['maf']>=5.0], on=dumascol, how='outer', right_index=True, left_index=True)
			else:
				merged = book.BP35[i][ book.BP35[i]['maf']>= 5.0 ]
		merged = merged[book.col_DumaPosition].values.tolist()
		for i in range(book.nsheets):
			if P0 is not None:
				P0 = merge(P0, book.BP35[i][ book.BP35[i][book.col_DumaPosition].isin(merged) ], on=dumascol, how='outer',
					left_index=True, right_index=True)
			else:
				P0 = book.BP35[i][ book.BP35[i][book.col_DumaPosition].isin(merged) ]

		# P0.to_excel('a.xlsx')
		# bp = bp[ self.dumascol + []]
		rows = dataframe_to_rows(P0, index=False)

		for r_idx, row in enumerate(rows, 1):
			for c_idx, value in enumerate(row, 1):
				if str(value) != 'nan':
					ws.cell(row=r_idx, column=c_idx, value=value)

		
	def insert_value_in_cell(self,ws, book, rows, col, dcol, BP, minor_, colname, types):
		"""
			한 column에 데이터 삽입
			- rows : 데이터를 삽입할 row list
			- BP : 'sum' >= 35.0 처리된 DataFrame
			- minor_ : s1~s2
			- maf : s / l
			- types : GPS or MAF
			리팩토링 필요, GPS 와 maf 한번에 처리
		"""
		index = book.GenomeStructure if colname == book.col_GenomeStructure else book.RepeatRegion
		# For Minor Exception
		if len(BP) == 0:
			for i in range(0, len(rows)):
				ws[col+str(rows[i])] = 0 if types == "GPS" else '-'
			ws[col+str(rows[len(rows)-1] + 1)] = 0 if types == "GPS" else '-'
		else:
			# merged = self.merge_Genome_Structure(Passage, minor_, sum_, colname)
			# _, tm_minor = self.get_level_major_minor(minor_, minor_, sum_, s1,s2)
			mid_rows = BP.loc[ minor_.index ]
			total_, cnt_sum=0, 0

			for i in range(0, len(rows)): # rows[i], index[i]		
				# 각 index 별로 그룹핑
				mrows = mid_rows [ mid_rows[ colname ] == index[i] ]
				if types == "GPS":
					ws[col+str(rows[i])] = len(mrows)
					cnt_sum+=len(mrows)

					# 180612 추가, 3개 컬럼, 강제 지정
					ws[dcol+str(rows[i])].value = ws[dcol+str(rows[i])].value + len(mrows) if ws[dcol+str(rows[i])].value != None else len(mrows)
				elif types == "MAF":
					if len(mrows) != 0:
						maf_ = divide(mrows[["minor"]], mrows[["sum"]]) * 100
						cnt_sum += len(maf_)
						total_ += maf_.sum()[0]
						ws[col+str(rows[i])] = str(round(maf_.sum()[0] / len(maf_),3)) + '%' if len(maf_) is not 0 else '-'
						
						# 180612 추가, 3개 컬럼, 강제 지정
						ws[dcol+str(rows[i])].value = float(ws[dcol+str(rows[i])].value) + maf_.sum()[0] if ws[dcol+str(rows[i])].value != None else maf_.sum()[0]
					else:
						ws[col+str(rows[i])] = '-'
						ws[dcol+str(rows[i])].value = ws[dcol+str(rows[i])].value if ws[dcol+str(rows[i])].value != None else 0
			if types == "GPS":
				ws[col+str(rows[len(rows)-1] + 1)] = cnt_sum
				# 180612 추가, 3개 컬럼, 강제 지정
				ws[dcol+str(rows[len(rows)-1] + 1)] = float(ws[dcol+str(rows[len(rows)-1] + 1)].value) + cnt_sum if ws[dcol+str(rows[len(rows)-1] + 1)].value != None else cnt_sum
			elif types == "MAF":
				# s, l = self.book.get_Number_of_GPS(BP, s1, s2)
				ws[col+str(rows[len(rows)-1] + 1)] = str(round(total_ / cnt_sum, 3)) + '%' if cnt_sum is not 0 else '-'
				# 180612 추가, 3개 컬럼, 강제 지정
				ws[dcol+str(rows[len(rows)-1] + 1)] = float(ws[dcol+str(rows[len(rows)-1] + 1)].value) + total_ if ws[dcol+str(rows[len(rows)-1] + 1)].value != None else total_

	def sheet1(self,ws, book):
		infolog("Start Sheet1")

		ws.title = "Polymorphic site"
		ws['A1'] = "strain"
		ws['C1'] = "Genome length (bp)"
		ws['D1'] = "Average of MAF"
		ws['E1'] = "Range of minor allele frequency(%)"
		ws.merge_cells(start_row=1, start_column=5, end_row=1, end_column=9)
		ws['E2'] = "2.5≤f<5"
		ws['F2'] = "5≤f<15"
		ws['G2'] = "15≤f<25"
		ws['H2'] = "25≤f"
		ws['I2'] = "sum"
		# A col
		ws.merge_cells(start_row=1, start_column=1, end_row=2, end_column=2)
		ws.merge_cells(start_row=1, start_column=3, end_row=2, end_column=3)
		ws.merge_cells(start_row=1, start_column=4, end_row=2, end_column=4)
		
		for i in range(book.nsheets):
			infolog("{0} Writing {1} sheet".format(time.time(), book.sheet_list[i]))
			
			ws['A' + str(i+3)] = book.sheet_list[i]
			ws.merge_cells(start_row=3+i, start_column=1, end_row=3+i, end_column=2)
			ws['C' + str(i+3)] = book.BPRawLength[i]
			s, l, _ = book.get_Number_of_GPS(book.BP35[i], 5.0, 51.0)
			ws['D' + str(i+3)] = str(round(s / l, 3)) + '%'
			ws['E' + str(i+3)] = book.get_Number_of_GPS(book.BP35[i], 2.5, 5.0)[1]
			ws["F" + str(i+3)] = book.get_Number_of_GPS(book.BP35[i], 5.0, 15.0)[1]
			ws["G" + str(i+3)] = book.get_Number_of_GPS(book.BP35[i], 15.0, 25.0)[1]
			ws["H" + str(i+3)] = book.get_Number_of_GPS(book.BP35[i], 25.0, 51.0)[1]
			ws["I" + str(i+3)] = book.get_Number_of_GPS(book.BP35[i], 5.0, 51.0)[1]

		infolog("End Sheet1")		

	def sheet2(self, ws, book):
		infolog("Start Sheet2")
		
		ws['A1'] = "strain"
		ws.merge_cells(start_row=1, start_column=1, end_row=2, end_column=2)
		ws['C1'] = "Genome length (bp)"
		ws.merge_cells(start_row=1, start_column=3, end_row=2, end_column=3)
		ws['D1'] = "Range of minor allele frequency(%)"
		ws.merge_cells(start_row=1, start_column=4, end_row=1, end_column=8)
		ws['D2'] = "2.5≤f<5"
		ws['E2'] = "5≤f<15"
		ws['F2'] = "15≤f<25"
		ws['G2'] = "25≤f"
		ws['H2'] = "sum"	

		for i in range(book.nsheets):
			infolog("Writing {0} sheet".format(book.sheet_list[i]))	

			ws['A' + str(i+3)] = book.sheet_list[i]
			ws['C' + str(i+3)] = book.BPRawLength[i]
			ws.merge_cells(start_row=3+i, start_column=1, end_row=3+i, end_column=2)
			
			col = 'D'
			for r1, r2 in zip(self.s1, self.s2):
				s, l, _ = book.get_Number_of_GPS(book.BP35[i], r1, r2) if len(book.BPxMinor[i])!=0 else (0,0)
				ws[col + str(i+3)] = str(round(s/l, 3))+'%' if l is not 0 else '-'
				col = next_col(col)
		infolog("End Sheet2")		

	def Full_Seq_in_sheet3(self, ws, col, book):
		infolog("Start Full_Seq_in_sheet3")
		
		# Writing list of GenomeStructure & RepeatRegion, NCR & ORF
		for r in range(4, 4+len(book.GenomeStructure)):
			ws[col+str(r)] = book.GenomeStructure[r-4]
		ws[col+str(4+len(book.GenomeStructure))] = 'Total'
		for r in range(0, len(book.RepeatRegion)):
			row = r+5+len(book.GenomeStructure)
			ws[col+str(row)] = book.RepeatRegion[r]
		ws[col+str(5+len(book.GenomeStructure)+len(book.RepeatRegion))] = 'Total'

		rs_ = 5+len(book.GenomeStructure)+len(book.RepeatRegion)
		ws[col+str(rs_+1)], ws[col+str(rs_+2)] = 'ORF', 'NCR'

		# Writing the count of each sheets
		col = next_col(col)
		for i in range(0, book.nsheets):
			infolog("{0} Writing {1} sheet".format(time.time(), i))
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
			col = next_col(col)
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
			col = next_col(col)
		infolog("End Full_Seq_in_sheet3")

	def sheet3(self,ws, book):
		infolog("Start sheet3")
		
		for i in range(0, len(self.s1)):
			infolog("Writing {0} ~ {1}".format(self.s1[i], self.s2[i]))

			# r : (s1,s2)의 범위 별 데이터의 row 변수
			r = i*(15+len(book.GenomeStructure) + len(book.RepeatRegion))+ 1
			ws['A'+str(r)] = "Region"
			ws.merge_cells(start_row=r, start_column=1, end_row=r+2, end_column=2)
			ws['C'+str(r)] = "Dumas Length (bp)"
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
				nlen = len(book.Dumas[ book.Dumas[ book.col_GenomeStructure ] == book.GenomeStructure[ix]]) # dumas length 기준
				ws['C' + str(G_rows[ix])] = nlen
				cnt_num += nlen
			ws['B' + str(r+3+len(book.GenomeStructure))] = "Total"
			ws['C' + str(r+3+len(book.GenomeStructure))] = cnt_num
			cnt_num = 0

			nr = r+4+len(book.GenomeStructure)+len(book.RepeatRegion)
			R_rows = list(range(r+4+len(book.GenomeStructure), nr))
			for ix in range(0, len(book.RepeatRegion)):
				ws['B' + str(R_rows[ix])] = book.RepeatRegion[ix]
				nlen = len(book.Dumas[ book.Dumas[ book.col_RepeatRegion ] == book.RepeatRegion[ix]])
				ws['C'+str(R_rows[ix])] = nlen
				cnt_num += nlen
			ws['B'+str(nr)] = 'Total'
			ws['C' + str(nr)] = cnt_num

			## ORF NCR
			ws["A"+str(nr+1)] = "ORF"
			ws["A"+str(nr+2)] = "NCR"
			# ORF NCR Full Length
			ws['C'+str(nr+1)] = len(book.Dumas[ book.Dumas[book.col_ORF].isin(book.ORF)])
			ws['C'+str(nr+2)] = len(book.Dumas[ book.Dumas[book.col_ORF].isin(book.NCR)])
			# end B, C columns
			
			# ncol = 'D'
			# ncol : Column of 'Number of GPS' is need when calculating 'Number of GPS / Length'
			ncol = col = 'D'

			ws[col+str(r)] = "Number of GPS"
			gcol = next_n_col(col, book.nsheets)
			for c in range(0, book.nsheets):
				ws[col+str(r+2)] = book.sheet_list[c]
				# 35 이상 
				s, l, minor_ = book.get_Number_of_GPS(book.BP35[c], self.s1[i], self.s2[i])
				# maf_ = str(round(s/l, 3)) + '%' if l is not 0 else '-'

				tx_rows = book.BP35[c].loc[ minor_.index ]
				tx_orf = tx_rows[ tx_rows[ book.col_ORF ].isin(book.ORF)]
				tx_ncr = tx_rows[ tx_rows[ book.col_ORF ].isin(book.NCR)]

				ws.merge_cells(start_row=r, start_column=4, end_row=r, end_column=4 + book.nsheets)
				ws[col+str(r+1)] = book.filename
				ws.merge_cells(start_row=r+1, start_column=4, end_row=r+1, end_column=4 + book.nsheets)
				self.insert_value_in_cell(ws, book, G_rows, col, gcol, book.BP35[c], minor_, book.col_GenomeStructure , "GPS")
				self.insert_value_in_cell(ws, book, R_rows, col, gcol, book.BP35[c], minor_, book.col_RepeatRegion , "GPS")
				ws[col+str(nr+1)] = len(tx_orf)
				ws[col+str(nr+2)] = len(tx_ncr)
				#180612
				ws[gcol+str(nr+1)] = float(ws[gcol+str(nr+1)].value) + len(tx_orf) if ws[gcol+str(nr+1)].value != None else len(tx_orf)
				ws[gcol+str(nr+2)] = float(ws[gcol+str(nr+2)].value) + len(tx_ncr) if ws[gcol+str(nr+2)].value != None else len(tx_ncr)
				col = next_col(col)
			col = next_col(col)

			# "Average MAF at GPS"
			ws[col+str(r)] = "Average MAF at GPS"
			dcol = next_n_col(col, book.nsheets)
			for c in range(0, book.nsheets):
				ws[col+str(r+2)] = book.sheet_list[c]
				ws.merge_cells(start_row=r, start_column=5 + book.nsheets, end_row=r, end_column=5 + 2 * book.nsheets)
				ws[col+str(r+1)] = book.filename
				ws.merge_cells(start_row=r+1, start_column=5 + book.nsheets, end_row=r+1, end_column=5 + 2 * book.nsheets)

				s, l, minor_ = book.get_Number_of_GPS(book.BP35[c], self.s1[i], self.s2[i])
				# maf_ = str(round(s/l, 3)) + '%' if l is not 0 else '-'
				
				tx_rows = book.BP35[c].loc[ minor_.index ]
				tx_orf = tx_rows[ tx_rows[ book.col_ORF ].isin(book.ORF)]
				tx_ncr = tx_rows[ tx_rows[ book.col_ORF ].isin(book.NCR)]

				self.insert_value_in_cell(ws, book, G_rows, col, dcol, book.BP35[c], minor_, book.col_GenomeStructure , "MAF")
				self.insert_value_in_cell(ws, book, R_rows, col, dcol, book.BP35[c], minor_, book.col_RepeatRegion , "MAF")

				# 180612
				maf_ = (divide(tx_orf[["minor"]], book.BP35[c][['sum']].loc[ tx_orf.index]) * 100).sum()[0]
				ws[col+str(nr+1)] = str(round( maf_ /len(tx_orf), 3)) +'%' if len(tx_orf) is not 0 else '-'
				ws[dcol+str(nr+1)] = float(ws[dcol+str(nr+1)].value) + maf_ if ws[dcol+str(nr+1)].value != None else maf_

				maf_ = (divide(tx_ncr[["minor"]], book.BP35[c][['sum']].loc[ tx_ncr.index]) * 100).sum()[0]
				ws[col+str(nr+2)] = str(round( maf_ /len(tx_ncr), 3)) +'%' if len(tx_ncr) is not 0 else '-'
				ws[dcol+str(nr+2)] = float(ws[dcol+str(nr+2)].value) + maf_ if ws[dcol+str(nr+2)].value != None else maf_

				col = next_col(col)

			col = next_col(col)
			# "Number of GPS / length"
			ws[col+str(r)] = "Number of GPS / length * 100"
			d2col = next_n_col(col, book.nsheets)
			for c in range(0, book.nsheets):
				ws[col+str(r+2)] = book.sheet_list[c]
				ws.merge_cells(start_row=r, start_column=6 + 2 *book.nsheets, end_row=r, end_column=6 + 3 *book.nsheets)
				ws[col+str(r+1)] = book.filename
				ws.merge_cells(start_row=r+1, start_column=6 + 2 *book.nsheets, end_row=r+1, end_column=6 + 3 *book.nsheets)
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

					ws[col + str(rs)] = (ws[ncol+str(rs)].value / llen_) * 100 if llen_ is not 0 else '-'
					ridx = ridx + 1

					# 180612 강제지정
					ws[d2col + str(rs)] = ws[d2col + str(rs)].value + llen_ if ws[d2col + str(rs)].value != None else llen_

				# ORF & NCR
				for rs in range(2):
					tmp, rs_ = 0, r+5+len(book.GenomeStructure)+len(book.RepeatRegion) + rs
					if rs == 0:
						tmp = len(book.BP35[c][ book.BP35[c][book.col_ORF].isin(book.ORF)])
						ws[col+str(rs_)] = (ws[ncol+str(rs_)].value / tmp) * 100 if tmp != 0 else '-'
						
					else:
						tmp = len(book.BP35[c][ book.BP35[c][book.col_ORF].isin(book.NCR)])
						ws[col+str(rs_)] = (ws[ncol+str(rs_)].value / tmp) * 100 if tmp != 0 else '-'

					# 180612 강제지정
					ws[d2col + str(rs_)] = ws[d2col + str(rs_)].value + tmp if ws[d2col + str(rs_)].value != None else tmp

				ncol = next_col(ncol)
				col = next_col(col)				

			# 180612 강제지정 
			for rs in range(r+3, r+7+len(book.GenomeStructure)+len(book.RepeatRegion)):
				ws[gcol+str(rs)].value = float(ws[gcol+str(rs)].value) / book.nsheets
				ws[dcol+str(rs)] = str(round(float(ws[dcol+str(rs)].value) / float(ws[gcol+str(rs)].value), 3)) + '%' if float(ws[dcol+str(rs)].value) != 0 else '-'
				ws[d2col + str(rs)] = (float(ws[d2col + str(rs)].value) / ( float(ws[d2col + str(rs)].value) / 3)) * 100 if float(ws[d2col + str(rs)].value) != 0.0 else 0

			col = next_col(col)
			ws[col+str(r+3)] = str(self.s1[i]) + "~" + str(self.s2[i]) + "%" if self.s2[i] != 51.0 else str(self.s1[i]) + "% 이상"			
		infolog("finished writing GenomeStructure & RepeatRegion")

		#### Full Sequence				
		col = next_col(next_col(next_col(col)))
		self.Full_Seq_in_sheet3(ws, col, book)
		infolog("End sheet3")

	def sheet4_5(self, ws, title, book):
		infolog("Start sheet {0}".format(4 if title is "ORF" else 5))
		ws['B2'] = book.filename
		ws["C2"] = "(BP_full)Length"
		ws.merge_cells(start_row=2, start_column=3, end_row=2, end_column=3+book.nsheets-1)
		ws["B3"] = title
		col = 'C'
		for i in range(0, book.nsheets):
			ws[col + str(3)] = book.sheet_list[i]
			col = next_col(col)		

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
			col = next_col(col)

		# s1~s2 범위
		col = chr(ord('A')+3+book.nsheets)
		for i in range(0, len(self.s1)):
			co = (5+3*book.nsheets) * i + (5+book.nsheets)
			ws[col+str(2)] = str(self.s1[i]) + '~' + str(self.s2[i]) +'%' if self.s2[i] != 51.0 else str(self.s1[i]) + '%~'
			ws[col+str(3)] = title # NCR or ORF
			for ix in range(0, len(orf_)):
				ws[col+str(rows[ix])] = orf_[ix]
			ws[col+str(len(orf_) + 4)] = 'total'

			# Number of GPS of each sheets
			gps = list()
			col = next_col(col)			
			ws[col+str(2)] = "Number of GPS"
			ws.merge_cells(start_row=2, start_column=co, end_row=2, end_column=co+book.nsheets)

			dcol = next_n_col(col, book.nsheets)
			for n in range(0, book.nsheets):
				ws[col+str(3)] = book.sheet_list[n]
				_, _, tx_minor = book.get_Number_of_GPS(book.BP35[n], self.s1[i], self.s2[i])
				tx_rows = book.BP35[n].loc[ tx_minor.index ]

				g = []
				cnt_ = 0
				for ix in range(0, len(orf_)):
					mrows = tx_rows[ tx_rows[book.col_ORF] == orf_[ix] ]
					ws[col+str(rows[ix])] = len(mrows)
					cnt_ += len(mrows)
					g.append(len(mrows))

					# 180612
					ws[dcol+str(rows[ix])] = float(ws[dcol+str(rows[ix])].value) + len(mrows) if ws[dcol+str(rows[ix])].value != None else len(mrows)

				ws[col+str(len(orf_) + 4)] = cnt_
				ws[dcol+str(len(orf_) + 4)] = float(ws[dcol+str(len(orf_) + 4)].value) + cnt_ if ws[dcol+str(len(orf_) + 4)].value != None else cnt_
				g.append(cnt_)
				gps.append(g)
				col = next_col(col)
			col = next_col(col)

			gcol = next_n_col(col, book.nsheets)
			# Average MAF at GPS  of each sheets
			ws[col+str(2)] = "Average MAF at GPS"
			ws.merge_cells(start_row=2, start_column=co+book.nsheets+1, end_row=2, end_column=co+2*book.nsheets+1)
			for n in range(0, book.nsheets):
				val, cnt_ = 0, 0
				ws[col+str(3)] = book.sheet_list[n]
				_, _, tx_minor = book.get_Number_of_GPS(book.BP35[n], self.s1[i], self.s2[i])
				tx_rows = book.BP35[n].loc[ tx_minor.index ]
				for ix in range(0, len(orf_)):
					mrows = tx_rows[ tx_rows[book.col_ORF] == orf_[ix] ]
					if len(mrows) is 0:
						ws[col+str(rows[ix])] = '-'

						# 180612
						ws[gcol+str(rows[ix])] = ws[gcol+str(rows[ix])].value if ws[gcol+str(rows[ix])].value != None else 0
					else:
						maf_ = divide(mrows[["minor"]], mrows[["sum"]]) * 100
						val += maf_.sum()[0]
						cnt_ += len(maf_)
						ws[col+str(rows[ix])] = str(round(maf_.sum()[0] / len(maf_), 3)) + '%'

						# 180612
						ws[gcol+str(rows[ix])] = float(ws[gcol+str(rows[ix])].value) + maf_.sum()[0] if ws[gcol+str(rows[ix])].value != None else maf_.sum()[0]
				# total
				ws[col+str(len(orf_)+4)] = str(round(val / cnt_, 3)) + '%' if cnt_ is not 0 else '-'
				ws[gcol+str(len(orf_)+4)] = float(ws[gcol+str(len(orf_)+4)].value) + val if ws[gcol+str(len(orf_)+4)].value != None else val
				col = next_col(col)
			col = next_col(col)
			# "Number of GPS / Length" of each sheets
			ws[col+str(2)] = "Number of GPS / Length * 100"
			ws.merge_cells(start_row=2, start_column=co+2*book.nsheets+2, end_row=2, end_column=co+3*book.nsheets+2)
			ncol = 'C'
			d2col = next_n_col(col, book.nsheets)
			for n in range(0, book.nsheets):
				ws[col+str(3)] = book.sheet_list[n]
				for ix in range(0, len(orf_)):
					l_ = int(str(ws[ncol+str(rows[ix])].value))
					ws[col+str(rows[ix])] = (gps[n][ix] / l_)*100 if l_ is not 0 else '-'
					ws[d2col+str(rows[ix])] = float(ws[d2col+str(rows[ix])].value) + l_ if ws[d2col+str(rows[ix])].value != None else l_
				# total
				l_ = int(str(ws[ncol+str(len(orf_)+4)].value))
				ws[col+str(len(orf_)+4)] = (gps[n][-1] / l_)*100 if l_ is not 0 else '-'
				ws[d2col+str(len(orf_)+4)] = float(ws[d2col+str(len(orf_)+4)].value) + l_ if ws[d2col+str(len(orf_)+4)].value != None else l_

				ncol = next_col(ncol)
				col = next_col(col)
			
			for row_ in range(4, len(orf_)+5):
				ws[dcol+str(row_)] = float(ws[dcol+str(row_)].value) / book.nsheets if float(ws[dcol+str(row_)].value) != 0 else 0
				ws[gcol+str(row_)] = str(round(float(ws[gcol+str(row_)].value) / float(ws[dcol+str(row_)].value), 3)) + '%' if float(ws[dcol+str(row_)].value) != 0 else '-'
				ws[d2col+str(row_)] = float(ws[dcol+str(row_)].value) / (float(ws[d2col+str(row_)].value) / 3) * 100 if float(ws[d2col+str(row_)].value) != 0 else 0

			# end for s1
			col = next_col(next_col(col))
		infolog("End sheet {0}".format( 4 if title is "ORF" else 5))

	def sheet6(self,ws, book):
		infolog("Start sheet6")
		col = 'B'
		for i in range(0, len(self.s1)):
			ws[col+'2'] = str(self.s1[i]) + '~' + str(self.s2[i]) +'%' if self.s2[i] != 51.0 else str(self.s1[i]) + '%~'

			# get basepair between s1 and s2
			mxr = []
			for n in range(0, book.nsheets):
				_, _, minor_ = book.get_Number_of_GPS(book.BP35[n], self.s1[i], self.s2[i])
				mxr.append(book.BP35[n].loc[minor_.index])

			for x in range(-1, book.nsheets):
				ws[col+'3'] = book.sheet_list[x] if x != -1 else "Major/Minor"
				rows = 4
				for a in range(0, len(book.col_basepair)):
					for b in range(0, len(book.col_basepair)):
						if a==b:
							continue
						ws[col+str(rows)] = len(mxr[x][logical_and( mxr[x]['major_idx']==a, 
							mxr[x]['minor_idx']==b )]) if x != -1 else book.col_basepair[a] + '/' + book.col_basepair[b]
						rows = rows + 1
				col = next_col(col)
			col = next_col(col)
		infolog("End sheet6")

	def sheet7(self, ws, book):
		infolog("Start sheet7")
		for i in range(0, len(self.s1)):
			r = i * (book.nsheets + 12) + 2
			ws["A"+str(r)] = str(self.s1[i]) + "~" + str(self.s2[i]) + "%" if self.s2[i] != 51.0 else str(self.s1[i]) + "%~"
			
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

			ws['E'+str(r+3+book.nsheets)] = 'sum'
			for x in range(0, book.nsheets):
				row, mafrow = r+x+3, r+book.nsheets+x+5
				s, l, minor_ = book.get_Number_of_GPS(book.BP35[x], self.s1[i], self.s2[i])
				mxr = book.BP35[x].loc[minor_.index]
				# s, l = self.book.get_Number_of_GPS(self.PxMinor[x], self.PxSum[x], self.s1[i], self.s2[i]) if len(self.PxMinor[x])!=0 else (0,0)
				ws["B"+str(row)] = book.sheet_list[x]
				ws['C'+str(row)] = len(minor_)
				ws['D'+str(row)] = str(round(s/l, 3)) + '%' if l != 0 else '-'
				
				col = 'F'
				for a in range(0, 4):
					for b in range(0, 4):
						if a == b:
							continue
						else:
							ws[col+str(r+2)] = book.col_basepair[b].lower()
							ext_ = mxr[logical_and( mxr['major_idx']==a, mxr['minor_idx']==b )]

							ws[col+str(row)] = len(ext_)

							if len(ext_) is not 0:
								# maf들의 평균
								maf_ =  (divide( ext_[['minor']], ext_[['sum']]) * 100).sum()[0]						
								ws[col+str(mafrow)] = str(round( maf_ / len(ext_), 3)) + '%'
							else:
								ws[col+str(mafrow)] = 'N/A'

							# 전체 개수의 maf
							if ws[col+str(r+3+book.nsheets)].value == None:							
								ws[col+str(r+3+book.nsheets)] =  len(ext_)
							else:
								ws[col+str(r+3+book.nsheets)].value += len(ext_)

							if ws[col+str(r+5+ 2*book.nsheets)].value == None:
								ws[col+str(r+5+ 2*book.nsheets)] = maf_
								# ws[col+str(r+5+ 2*book.nsheets+1)] = ext_[['sum']].sum()[0]
							else:
								# print(i, "mj ", a, "mn ", b, ext_[['minor']].sum()[0])
								ws[col+str(r+5+ 2*book.nsheets)].value += maf_
								# ws[col+str(r+5+ 2*book.nsheets+1)].value += ext_[['sum']].sum()[0]


						col = next_col(col)
			# total maf
			col = 'F'
			numrow, mnrow = r+3+book.nsheets, r+5+2*book.nsheets
			for c in range(0, 12):
				ws[col+str(mnrow)] = str(round( float(ws[col+str(mnrow)].value) / float(ws[col+str(numrow)].value) * 100, 3))+'%' if ws[col+str(numrow)].value != 0 else 'N/A'
				col = next_col(col)

		infolog("End sheet7")
