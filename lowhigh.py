"""
 Copyright [2017] [Il Seop Lee / Chungbuk National University]
 
 Licensed under the Apache License, Version 2.0 (the "License");
 you may not use this file except in compliance with the License.
 You may obtain a copy of the License at
 
 http://www.apache.org/licenses/LICENSE-2.0
 
 Unless required by applicable law or agreed to in writing, software
 distributed under the License is distributed on an "AS IS" BASIS,
 WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 See the License for the specific language governing permissions and
 limitations under the License.
 """
from analyzer import infolog, next_col #Analyzer.get_Number_of_GPS, next_col, infolog
from numpy import logical_and, logical_or
import pandas as pd
from numpy import divide, zeros, ones

class LowHigh:
	"""
		analysis of increase or decrease sequence
	"""
	def __init__(self, book, isAverage=False, isCartesian=False):
		self.BPmergedinc = []

		self.BPmergeddec = []

		self.s1 = book.s1 #[5.0, 15.0, 25.0, 5.0]

		self.s2 = book.s2 #[15.0, 25.0, 51.0, 51.0]

		self.dumascol = [book.col_GenomeStructure, book.col_RepeatRegion, book.col_ORF, book.col_DumaPosition, book.col_DumaSeq]
		if not isAverage and not isCartesian:
			self.Init_Individual(book)
		elif isAverage:
			self.Init_Average(book)
		elif isCartesian:
			self.Init_Cartesian(book)
		
	# todo : 테스트
	def Init_Average(self, book):
		# book.save_data(book, 'Average_book_vaccine.pkl')

		for sn in book.sheet_list:
			assert sn[0] == 'A' or sn[0] == 'B', "sheet name"

		# for i in range(0,len(book.nsheets)):
		# 	book.BPRaw[i]['sum_'+ book.sheet_list[i]] = book.BPRaw[i][book.basepair].sum(axis=1)

		# A & B 위치별 MAF 평균
		# Avg(A) ~ Avg(B)의 변화 분석
		infolog("Init_Average")

		# # 필요 최종 column : dumacol(5), diffofdec, diffofinc, minor_idx_x, minor_idx_y, ->major(2) / minor..
		# # 필요 값 : any maf 5% 이상, all maf 5% 이상, 값이 있는 곳

		## 정한 값 :	minor_idx_x -> A의 가장 마지막.. / minor_idx_y -> B의 가장 마지막
		## 			minor_x -> A의 minor 평균 / minor_y -> B의 minor 평균
		##			
		for j, i in enumerate(book.BP35):
			i['maf_'+book.sheet_list[j]] = divide(i[['minor']], i[['sum']]) * 100
			i[book.sheet_list[j][0].upper()+'_minor_idx'] = i[['minor_idx']]
			book.BP35[j] = i.drop(columns='minor_idx')

		merged = book.BP35[0].drop(columns=['seq', 'sum', 'A', 'G', 'C', 'T', 'major'])
		merged['maf_'+book.sheet_list[0][0].upper()] = book.BP35[0]['maf_'+book.sheet_list[0]]
		for i in range(1, book.nsheets):
			book.BP35[i] = book.BP35[i].drop(columns=['seq', 'sum', 'A', 'G', 'C', 'T', 'major'])
			grp = book.sheet_list[i][0].upper()
			col_maf = 'maf_' + grp
			col_minor_idx = grp+'_minor_idx'

			merged = pd.merge(merged, book.BP35[i], on=self.dumascol, how='outer')
			if not merged.columns.contains(col_maf): 
				merged[col_maf] = merged['maf_' + book.sheet_list[i] ]
			else : 
				merged[col_maf] = merged[[col_maf, 'maf_'+book.sheet_list[i]]].sum(axis=1)

			if merged.columns.contains(col_minor_idx+'_x'):
			    merged[col_minor_idx] = merged[col_minor_idx+'_y']
			    merged = merged.drop(columns=col_minor_idx+'_x').rename(index=str, columns={col_minor_idx+'_y' : col_minor_idx})

			merged = merged[ merged['major_idx_x'] == merged['major_idx_y'] ]
			merged = merged.drop(columns = 'major_idx_x').rename(index=str, columns={"major_idx_y": "major_idx"})

		Acol, Bcol = [], []
		for i in range(len(book.sheet_list)):
			if book.sheet_list[i][0].upper() == 'A':
				Acol.append('maf_'+book.sheet_list[i])
			else:
				Bcol.append('maf_'+book.sheet_list[i])

		merged['maf_A'] = merged['maf_A'] / len(Acol)
		merged['maf_B'] = merged['maf_B'] / len(Bcol)

		A_all, A_any = True, False
		B_all, B_any = True, False
		for i in Acol:
			A_all = logical_and( A_all, merged[i] >= 5.0 )
			A_any = logical_or( A_any, merged[i] >= 5.0 )
		    
		for i in Bcol:
			B_all = logical_and( B_all, merged[i] >= 5.0 )
			B_any = logical_or( B_any, merged[i] >= 5.0 )
		# A, B = [], []

		merged['minor_idx_x'] = merged['A_minor_idx'].T[-1:].T
		merged['minor_idx_y'] = merged['B_minor_idx'].T[-1:].T

		merged['major_idx_x'] = merged['major_idx']
		merged['major_idx_y'] = merged['major_idx']
		merged = merged.drop(columns = ['A_minor_idx', 'B_minor_idx'])

		tmpx = merged['minor_x'].sum(axis=1) / len(Acol)
		tmpy = merged['minor_y'].sum(axis=1) / len(Bcol)
		merged = merged.drop(columns=['minor_x', 'minor_y'])
		merged['minor_x'] = tmpx
		merged['minor_y'] = tmpy

		merged['diffofinc'] = merged['maf_B'] - merged['maf_A']
		merged['diffofdec'] = merged['maf_A'] - merged['maf_B']

		cond2 = merged['diffofinc'] >= 5.0
		cond3 = merged['diffofdec'] >= 5.0

		maf_5_all = merged[ logical_and(A_all, B_all).values ]
		maf_5_any = merged[ logical_or(A_any, B_any).values ]

		merged.to_excel('{}_merged.xlsx'.format(book.filename))
		maf_5_all.to_excel('{}_maf_5_all.xlsx'.format(book.filename))
		maf_5_any.to_excel('{}_maf_5_any.xlsx'.format(book.filename))

		# print(len(merged), len(maf_5_all), len(maf_5_any))

		self.BPmergedinc.append( merged)
		self.BPmergedinc.append(maf_5_all)
		self.BPmergedinc.append(maf_5_any)

		self.BPmergeddec.append( merged)
		self.BPmergeddec.append(maf_5_all)
		self.BPmergeddec.append(maf_5_any)
		book.sheet_list = ['All', 'All_5', 'Any_5']
		book.nsheets = len(book.sheet_list)
		infolog("End Init_Average")

	def Init_Cartesian(self, book):
		for sname in book.sheet_list:
			assert sname[0] == 'A' or sname[0] == 'B', "sheet name"
		# book.save_data(book, 'Cartesian_book_vaccine.pkl')		
		Acol, Bcol = [], []
		for i, j in enumerate(book.BP35):
			grp = book.sheet_list[i][0].upper()
			j[grp+'_maf'] = divide( j[['minor']], j[['sum']]) * 100
			j[book.sheet_list[i]+'_maf'] = divide( j[['minor']], j[['sum']]) * 100
			j[grp+'_major_idx'], j[grp+'_minor_idx'] = j[['major_idx']], j[['minor_idx']]
			j[grp+'_minor'] = j[['minor']]
			j = j.drop(columns=['minor', 'major_idx', 'minor_idx'])

			if grp == 'A': Acol.append(j)
			else: Bcol.append(j)

		merged = None
		for A in Acol:
			if merged is None: merged = A
			else: merged = pd.merge(merged, A, on=self.dumascol, how='outer')
			merged = merged.drop(columns=['sum', 'A', 'G', 'C', 'T', 'major'])
			for B in Bcol:
				merged = pd.merge(merged, B, on=self.dumascol, how='outer')
				merged = merged[ merged['A_major_idx'] == merged['B_major_idx'] ]
				if merged.columns.contains('maf'): merged['maf'] = merged['maf'] + abs(merged['A_maf'] - merged['B_maf'])
				else: merged['maf'] = abs(merged['A_maf'] - merged['B_maf'])

				merged = merged.drop(columns = ['B_maf', 'B_major_idx', 'sum', 'A', 'G', 'C', 'T', 'major'])
			merged = merged.drop(columns=['A_maf', 'A_major_idx'])

		_all, _any = True, False
		for sn in book.sheet_list:
			_all = logical_and( _all, merged[sn+'_maf'] >= 5.0)
			_any = logical_or( _any, merged[sn+'_maf'] >= 5.0)

		merged['major_idx_x'], merged['major_idx_y'] = True, True
		merged['diffofinc'] = merged['maf'] / float(len(Acol) * len(Bcol))
		maf_5_all = merged[ _all.values ]
		maf_5_any = merged[ _any.values ]

		merged.to_excel('car_merged.xlsx')
		maf_5_all.to_excel('car_maf_5_all.xlsx')
		maf_5_any.to_excel('car_maf_5_any.xlsx')

		self.BPmergedinc.append(merged)
		self.BPmergedinc.append(maf_5_all)
		self.BPmergedinc.append(maf_5_any)

		book.sheet_list = ['All', 'All_5', 'Any_5']
		book.nsheets = len(book.sheet_list)

	def Init_Individual(self, book):
		infolog("lowhigh init start")
		sheet_names = []
		BPmaf = []
		for bp in book.BP35:
			#get maf 5%
			infolog("bp 5.0 ~ 50.0")
			s, l, minor_ = book.get_Number_of_GPS(bp, 5.0, 51.0)
			BPmaf.append(bp.loc[minor_.index])

		
		for i, x in enumerate(BPmaf):
			for j, y in enumerate(BPmaf):
				if i >= j:
					continue
				infolog("{0} -> {1}".format(i,j))
				sheet_names.append(book.sheet_list[i]+'->'+book.sheet_list[j])
				# low의 maf5%이상과 high의 maf5%이상 값을 병합
				# 둘 중 하나라도 sum >= 35.0 and maf 5% 이상인 값에서 분석을 진행
				dumaspos = pd.merge(x,y, how='outer', on=self.dumascol, suffixes=('_x', '_y'), right_index=True, left_index=True)[[book.col_DumaPosition]]

				# dumas position 기준으로 양쪽 위치 추출
				x1 = book.BP35[i][ book.BP35[i][book.col_DumaPosition].isin(dumaspos[book.col_DumaPosition].values.tolist())] if len(dumaspos) is not 0 else pd.DataFrame()
				y1 = book.BP35[j][ book.BP35[j][book.col_DumaPosition].isin(dumaspos[book.col_DumaPosition].values.tolist())] if len(dumaspos) is not 0 else pd.DataFrame()

				merged = pd.merge(x1,y1, how='outer', on=self.dumascol, suffixes=('_x', '_y'), right_index=True, left_index=True)
				
				if(len(x1) is 0 or len(y1) is 0):
					self.BPmergedinc.append(pd.DataFrame())
					self.BPmergeddec.append(pd.DataFrame())
					continue
								
				x_maf = divide(x1[['minor']], x1[['sum']]) * 100
				y_maf = divide(y1[['minor']], y1[['sum']]) * 100
				
				# 증가하는 위치와 감소하는 위치
				## 양쪽에 서로 없는 index가 있으면 Error - 한쪽에 값이 없으면 NaN : dropna함수로 제거
				merged['diffofinc'] = y_maf - x_maf
				merged['diffofdec'] = x_maf - y_maf

				cond1 = merged['major_idx_x'] == merged['major_idx_y']
				cond2 = merged['diffofinc'] >= 5.0
				cond3 = merged['diffofdec'] >= 5.0				

				# 한쪽에 값이 없으면 drop되기 때문에 어디쪽에 해도 상관은 없음							
				self.BPmergedinc.append(pd.DataFrame.dropna(merged[ logical_and(cond1, cond2).values ], how='all'))
				self.BPmergeddec.append(pd.DataFrame.dropna(merged[ logical_and(cond1, cond3).values ], how='all'))

		book.sheet_list = sheet_names
		book.nsheets = len(sheet_names)
		infolog("lowhigh init end")

	def PolymorphicSite(self, ws, types_, book):
		infolog("lowhigh PolymorphicSite start")
		ws.title = "INC-Polymorphic site"
		# A col
		ws['A1'] = types_
		ws.merge_cells(start_row=1, start_column=1, end_row=2, end_column=2)
		
		# C col
		ws['C1'] = "Increase in genetic polymrphism (%)"
		ws.merge_cells(start_row=1, start_column=3, end_row=1, end_column=6)

		from numpy import logical_and
		for i in range(book.nsheets):
			ws['A' + str(i+3)] = book.sheet_list[i]
			ws.merge_cells(start_row=i+3, start_column=1, end_row=i+3, end_column=2)
			col = 'C'
			bp = self.BPmergedinc[i][['diffofinc']] if types_ == 'INC' else self.BPmergeddec[i][['diffofdec']]
			for s1_, s2_ in zip(self.s1, self.s2):
				ws[col+'2'] = str(s1_) + '≤n<' + str(s2_) # 'Sum' if s1_ == 5.0 and s2_ == 51.0 else str(s1_) + '≤n<' + str(s2_) 
				ws[col+str(i+3)] = len(bp[ logical_and(bp >= s1_, bp < s2_).values]) if len(bp) is not 0 else 0
				col = next_col(col)			
		infolog("lowhigh PolymorphicSite end")

	def GenomeStr(self, ws, types_, book):
		infolog("lowhigh GenomeStr start")
		bp = self.BPmergedinc if types_ == 'INC' else self.BPmergeddec		
		col_diff = "diffofinc" if types_ == 'INC' else "diffofdec"

		# s1 ~ s2 
		for i in range(len(self.s1)):
			rows = i*(14+len(book.GenomeStructure) + len(book.RepeatRegion))+ 1

			ws['A'+str(rows)] = 'Region'
			ws.merge_cells(start_row=rows, start_column=1, end_row=rows+1, end_column=1)
			ws['B'+str(rows)] = 'Dumas Length'
			ws.merge_cells(start_row=rows, start_column=2, end_row=rows+1, end_column=2)

			for r in range(2, 2+len(book.GenomeStructure)):
				ws['A' + str(rows+r)] = book.GenomeStructure[r-2]
				ws['B' + str(rows+r)] = len(book.Dumas[ book.Dumas[book.col_GenomeStructure] == book.GenomeStructure[r-2] ])

			ws['A' + str(rows+2+len(book.GenomeStructure))] = 'Total'
			ws['B' + str(rows+2+len(book.GenomeStructure))] = len(book.Dumas)

			sum_ = 0
			for r in range(len(book.RepeatRegion)):
				row = rows+r+3+len(book.GenomeStructure)
				ws['A' + str(row)] = book.RepeatRegion[r]
				cnt = len(book.Dumas[ book.Dumas[book.col_RepeatRegion] == book.RepeatRegion[r] ])
				ws['B' + str(row)] = cnt
				sum_ += cnt

			leng = rows+ len(book.GenomeStructure) + len(book.RepeatRegion)
			ws['A' + str(leng+3)] = 'Total'
			ws['B' + str(leng+3)] = sum_
			ws['A' + str(leng+4)] = 'ORF'
			ws['B' + str(leng+4)] = len(book.Dumas[ book.Dumas[book.col_ORF].isin(book.ORF)])	
			ws['A' + str(leng+5)] = 'NCR'
			ws['B' + str(leng+5)] = len(book.Dumas[ book.Dumas[book.col_ORF].isin(book.NCR)])

			col = 'C'
			ws[col+str(rows)] = '{0} in genetic polymorphism'.format("Increase" if types_ is 'INC' else "Decrease")
			ws.merge_cells(start_row=rows, start_column=3, end_row=rows, end_column=2+book.nsheets)
			ws[chr(ord(col)+book.nsheets)+str(rows)] = "Average {0}".format("Increase" if types_ is 'INC' else "Decrease")
			ws.merge_cells(start_row=rows, start_column=3+book.nsheets, end_row=rows, end_column=2+2*book.nsheets)
			for s in range(book.nsheets):
				idx = logical_and( bp[s][col_diff] >= self.s1[i], bp[s][col_diff] < self.s2[i] )
				bpx = bp[s][idx.values]
				
				ncol = chr(ord(col)+book.nsheets)
				ws[col+str(rows+1)] = book.sheet_list[s]
				ws[chr(ord(col)+book.nsheets)+str(rows+1)] = book.sheet_list[s]

				sum_, avg_ = 0, 0
				for r in range(2, 2+len(book.GenomeStructure)):
					cnt = bpx[ bpx[book.col_GenomeStructure] == book.GenomeStructure[r-2]][[col_diff]]
					ws[col+str(rows+r)] = len(cnt)
					ws[ncol+str(rows+r)] = str(round((cnt.sum()[0] / len(cnt)), 3))+'%' if len(cnt) is not 0 else 'N/A'
					avg_ += cnt.sum()[0]
					sum_ += len(cnt)
				# Total
				ws[col+str(rows+2+len(book.GenomeStructure))] = sum_
				ws[chr(ord(col)+book.nsheets)+str(rows+2+len(book.GenomeStructure))] = str(round(avg_/sum_, 3))+'%' if sum_ is not 0 else 'N/A'

				sum_, avg_ = 0, 0
				for r in range(len(book.RepeatRegion)):
					row = rows+r+3+len(book.GenomeStructure)
					cnt = bpx[ bpx[book.col_RepeatRegion] == book.RepeatRegion[r]][[col_diff]]
					ws[col+str(row)] = len(cnt)
					ws[ncol+str(row)] = str(round((cnt.sum()[0] / len(cnt)),3))+'%' if len(cnt) is not 0 else 'N/A'
					sum_ += len(cnt)
					avg_ += cnt.sum()[0]
				ws[col+str(3+leng)] = sum_
				ws[ncol+str(3+leng)] = str(round(avg_/sum_,3))+'%' if sum_ is not 0 else 'N/A'
				# ORF & NCR
				cnt = bpx[ bpx[book.col_ORF].isin(book.ORF)][[col_diff]]
				ws[col+str(4+leng)] = len(cnt)
				ws[ncol+str(4+leng)] = str(round(cnt.sum()[0] / len(cnt), 3))+'%' if len(cnt) is not 0 else 'N/A'
				cnt = bpx[ bpx[book.col_ORF].isin(book.NCR)][[col_diff]]
				ws[col+str(5+leng)] = len(cnt)
				ws[ncol+str(5+leng)] = str(round(cnt.sum()[0] / len(cnt), 3))+'%' if len(cnt) is not 0 else 'N/A'
				col = next_col(col)

			ncol = chr(ord(ncol)+book.nsheets)
			ws[ncol+str(rows+2)] = str(self.s1[i])+"~"+str(self.s2[i]) #if self.s2[i] != 51.0 else str(self.s1[i])+"~"
		infolog("lowhigh GenomeStr end")

	def ORFNCR(self, ws, types_, title, book):
		infolog("lowhigh {0} start".format(title))

		ws['A1'] = title
		ws.merge_cells(start_row=1, start_column=1, end_row=2, end_column=1)
		ws['B1'] = 'Length'
		ws.merge_cells(start_row=1, start_column=2, end_row=2, end_column=2)

		orfncr = book.ORF if title is 'ORF' else book.NCR
		col_diff = "diffofinc" if types_ == 'INC' else "diffofdec"
		bp = self.BPmergedinc if types_ == 'INC' else self.BPmergeddec

		sum_, col = 0, 'B'
		for si in range(-1, book.nsheets):
			sum_ = 0
			if si != -1:
				ws[col+'2'] = book.sheet_list[si] 
			for r in range(len(orfncr)):
				if si == -1:
					ws['A'+str(r+3)] = orfncr[r]
					cnt = len(book.Dumas[ book.Dumas[book.col_ORF] == orfncr[r] ])
					ws[col+str(r+3)] = cnt
					sum_ += cnt
				else:
					cnt = len( bp[si][ bp[si][book.col_ORF] == orfncr[r] ] )
					ws[col+str(r+3)] = cnt
					sum_ += cnt
			ws[col+str(3+len(orfncr))] = sum_
			col = next_col(col)
		ws['A'+str(3+len(orfncr))] = 'Total'

		col = next_col(col)
		for i in range(len(self.s1)):			
			# ORF or CNR 종류
			ws[col+'1'] = str(self.s1[i])+'~'+str(self.s2[i]) #if self.s2[i]!=51.0 else str(self.s1[i])+'~'
			for r in range(len(orfncr)):
				ws[col+str(r+3)] = orfncr[r]
			col = next_col(col)

			col_GPS = col
			ws[col+'1'] = '{0} in genetic polymorphism'.format("Increase" if types_ == 'INC' else "Decrease")
			for s in range(book.nsheets):
				ws[col+'2'] = book.sheet_list[s]
				bpx = bp[s][ logical_and(bp[s][col_diff] >= self.s1[i], bp[s][col_diff] < self.s2[i]).values ]

				sum_ = 0
				for r in range(3, 3+len(orfncr)):
					cnt = len(bpx[ bpx[book.col_ORF] == orfncr[r-3]])
					ws[col+str(r)] = cnt	
					sum_ += cnt					
				ws[col+str(3+len(orfncr))] = sum_
				col = next_col(col)			

			ws[col+'1'] = 'Average {0}'.format("Increase" if types_ == 'INC' else "Decrease")
			for s in range(book.nsheets):			
				ws[col+'2'] = book.sheet_list[s]
				bpx = bp[s][ logical_and(bp[s][col_diff] >= self.s1[i], bp[s][col_diff] < self.s2[i]).values ]
				sum_ = 0
				for r in range(3, 3+len(orfncr)):
					cnt = bpx[ bpx[book.col_ORF] == orfncr[r-3]][[col_diff]]					
					ws[col+str(r)] = str(round(cnt.sum()[0] / len(cnt), 3))+'%' if len(cnt) is not 0 else 'N/A'
					sum_ += cnt.sum()[0]
				col = next_col(col)

			ws[col+'1'] = 'Number of GPS / Length'
			for s in range(book.nsheets):			
				ws[col+'2'] = book.sheet_list[s]
				for r in range(3, 3+len(orfncr)):
					s, l = float(ws[col_GPS+str(r)].value), float(ws['B'+str(r)].value)
					ws[col+str(r)] = str( round(s/l, 6) ) if l != 0 else 'N/A'
				col_GPS = next_col(col_GPS)
				col = next_col(col)
			col = next_col(col)

		infolog("lowhigh {0} end".format(title))


	def BaseComp(self, ws, types_, book):
		"""
			major가 같은 것, minor도 같은 것만 추출
			minor x -> 증가 포함
			1. major_idx_x == major_idx_y - 치환 제거
			2. (minor_idx_x == minor_idx_y) or ((minor_idx_x != minor_idx_y) and minor_x == 0)
		"""
		infolog("lowhigh BaseComp start")
		ws.merge_cells(start_row=1, start_column=1, end_row=2, end_column=1)
		ws['B1'], ws['B2'] = 'Major', 'Minor'
		ws['D1'], ws['H1'], ws['L1'], ws['P1'] = 'A', 'G', 'C', 'T'
		ws['G2'], ws['K2'], ws['O2'] = 'a', 'a', 'a'
		ws['C2'], ws['L2'], ws['P2'] = 'g', 'g', 'g'
		ws['D2'], ws['H2'], ws['Q2'] = 'c', 'c', 'c'
		ws['E2'], ws['I2'], ws['M2'] = 't', 't', 't'
		cols = ['A','G','C','T']		

		bp = self.BPmergedinc if types_ is "INC" else self.BPmergeddec

		for r in range(book.nsheets):
			ws['A'+str(r+3)] = book.sheet_list[r]

		for r in range(book.nsheets):
			col = 'C'
			cond1 = bp[r]['major_idx_x'] == bp[r]['major_idx_y']
			cond2 = bp[r]['minor_idx_x'] == bp[r]['minor_idx_y']
			cond3 = logical_and( bp[r]['minor_idx_x'] != bp[r]['minor_idx_y'], bp[r]['minor_x'] == 0 )

			idx = logical_and(cond1, logical_or(cond2, cond3))
			tmp = bp[r][ idx.values ]

			for a in range(len(cols)):
				for b in range(len(cols)):
					if a == b:
						continue
					# major는 _x, _y 가 이미 같고, minor는 _y기준으로 하면 _x에서 minor가 0이던값 무시됨
					ws[col+str(r+3)] = len(tmp[ logical_and(tmp['major_idx_y'] == a, tmp['minor_idx_y'] == b).values ])
					col = next_col(col)
				col = next_col(col)

		infolog("lowhigh BaseComp end")

	def listofdata(self, ws, n_, types_, book):
		from analyzer import dataframe_to_rows
		# print(types_, type(types_))
		# infolog("writing list of {0} data".format(types_))		
		
		bp = self.BPmergedinc[n_] if types_ is "INC" else self.BPmergeddec[n_]
		# bp = bp[ self.dumascol + []]
		rows = dataframe_to_rows(bp, index=False)

		for r_idx, row in enumerate(rows, 1):
			for c_idx, value in enumerate(row, 1):
				if str(value) != 'nan':
					ws.cell(row=r_idx, column=c_idx, value=value)

