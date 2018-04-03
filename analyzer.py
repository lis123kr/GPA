from openpyxl import Workbook
import numpy as np
import pandas as pd
from logging import error, info
from datetime import datetime

def infolog(message):
	info("{0} {1}".format(datetime.now().isoformat(), message))
	print("{0} {1}".format(datetime.now().isoformat(), message))

def errorlog(message):
	error("{0} {1}".format(datetime.now().isoformat(), message))	

def next_col(asc):
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
		self.book.mergesequence()
		# 뒤로 미룹시다
		# get BP35, major, minor, index of major, minor
		self.book.preprocessing()

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
	# try:
		self.Analyze_type = Analyze_type
		if Analyze_type == "Difference_of_Minor":
			self.s1 = [ 0, 2.5, 5, 15, 25, 5]
			self.s2 = [ 2.5, 5, 15, 25, 51.0, 51.0]
			
			types_ = ["INC", "DEC"]
			from lowhigh import LowHigh
			lh = LowHigh(book)
			for i in range(0,2):
				infolog("{0} analysis".format(types_[i]))
				wb = Workbook()
				ws1 = wb.active				
				ws3 = wb.create_sheet("Genome structure")
				ws4 = wb.create_sheet("ORF")
				ws5 = wb.create_sheet("NCR")
				ws6 = wb.create_sheet("Base composition_1")
				
				lh.PolymorphicSite(ws1, types_[i], book)
				lh.GenomeStr(ws3,types_[i], book)
				lh.ORFNCR(ws4, types_[i], "ORF", book)
				lh.ORFNCR(ws5, types_[i], "NCR", book)
				lh.BaseComp(ws6, types_[i], book)
				wb.save("["+types_[i]+"분석]" + book.filename + '.xlsx')
		else: # full
			from fullseq import FullSeq
			fs = FullSeq(book)
			wb = Workbook()
			ws1 = wb.active
			ws2 = wb.create_sheet("GPS")
			ws3 = wb.create_sheet("Genome structure")
			ws4 = wb.create_sheet("ORF")
			ws5 = wb.create_sheet("NCR")
			ws6 = wb.create_sheet("Base composition_1")
			ws7 = wb.create_sheet("Base composition_2")
			# FullSeq.init_full(book)
			fs.sheet1(ws1, book)				
			fs.sheet2(ws2, book)			
			fs.sheet3(ws3, book)			
			fs.sheet4_5(ws4, "ORF", book)			
			fs.sheet4_5(ws5, "NCR", book)			
			fs.sheet6(ws6, book)			
			fs.sheet7(ws7, book)
			wb.save("[분석]" + book.filename + '.xlsx')
		book.P0.to_excel('[통합]'+book.filename+'.xlsx', index=False)
	# 	return "Success"
	# except PermissionError:
	# 	return "Permission"
	# except Exception as e:
	# 	logging.error("{0} {1}".format(datetime.now().isoformat(), e))
	# 	return "Error"





	
