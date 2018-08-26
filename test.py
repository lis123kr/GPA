import pickle as pk
from book import Book
from analyzer import Analyzer
# 0 : gps / 1 : 증감
flag = 0

pfile = 'Average_book_vaccine.pkl'

# input_files= [ 'genome.txt', 'repeat.txt', 'ORF.txt', 'NCR.txt']
# dir_ = '.\\test\\'
# inputs = {}
# for file in input_files:
# 	key = file.split('.')[0]
# 	file_ = dir_ + file
# 	with open(file_, 'r') as f:
# 		inputs[key] = f.read().replace('\n',',').strip(',').split(',')

book = Book.load_data(pfile)
Analyze_type = 'Average' #"Difference_of_Minor"
excel = Analyzer()
excel.Analyze(Analyze_type, book)
