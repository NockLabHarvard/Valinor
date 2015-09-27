# Script to accept as input two .xlsx files and 
# calculate the inter-rater kappa agreement values
#
# The input files should have only 1 sheet 
# The code column should be named 'Code'
# The instance number column should be labeled 'Segment'
#
# author: Karthik Dinakar
# affliation: MIT Media Lab & Nock Lab, Harvard
#
#Permission is hereby granted, free of charge, to any person obtaining a copy of this software and associated documentation files (the “Software”), to deal in the Software without restriction, including without limitation the rights to use, copy, modify, merge, publish, distribute, sublicense, and/or sell copies of the Software, and to permit persons to whom the Software is furnished to do so, subject to the following conditions:
#The above copyright notice and this permission notice shall be included in all copies or substantial portions of the Software.
#THE SOFTWARE IS PROVIDED “AS IS”, WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.

import csv
import os.path
from re              import findall, split
from csv             import reader
from subprocess      import call
from os.path         import basename
from sklearn.metrics import classification_report
from argparse        import ArgumentParser


# Function to define parent level codes
def getParentCodes():
	return ['BACKGROUND INFORMATION', 
	        'MENTAL HEALTH TREATMENT',
		'RISK FACTORS',
		'PSYCHIATRIC SYMPTOMS',
		'SELF-HARM THOUGHTS AND BEHAVIORS',
		'SOCIAL COMMUNICATION / POST INFORMATION',
		'SOCIAL MEDIA SPECIFIC']

#Function to parse an Excel file exported via MAXQDA annotation
def parseCSV(path):
	print path
	try:
		call( ["xlsx2csv", path, "tmp.csv"] )
		codes = {}
		with open("tmp.csv", "rU") as fp:
			rdr  = reader(fp)
			rows = [ row for row in rdr ]
			codeIndex = rows[0].index('Code')
			postIndex = rows[0].index('Segment')
			sets          = []
			errors        = []
			rows          = rows[1:]
			for i, row in enumerate(rows):
				if findall(r'\d+', row[postIndex]):
					sets.append([ findall(r'\d+', row[postIndex])[0], row[codeIndex] ])
				else:
					errors.append(row)
			if errors:
				for i in errors:
					for row in rows:
						if ( i[postIndex] in row[postIndex] ) and ( findall(r'\d+', row[postIndex]) ):
							sets.append( [ findall(r'\d+', row[postIndex])[0], i[codeIndex] ] )
							break
			for item in sets:
				codes[item[0]] = []
			for item in sets:
				codes[item[0]].append(item[1])
		call ( ["rm", "tmp.csv"] )
		return codes
	except Exception, e:
		print "Oops: " + str(e)

#Function to get metrics for annotation agreement
def getClassificationReport(gold, coder):
	y_gold  = []
	y_coder = []
	for code in gold:
		y_gold.append(gold[code])
		y_coder.append(coder[code])
	with open('tmp.csv','wb') as fp:
		print >> fp, classification_report(y_gold, y_coder)
	with open('tmp.csv','rb') as fp:
		reader = csv.reader(fp)
		scores = [ split("    *", ' '.join(row).replace(",","")) for row in reader if row ]
	parentCodes = getParentCodes()
	report = {}
	for parentCode in parentCodes:
		report[parentCode] = {'count': 0, 'score': [0.0, 0.0, 0.0, 0.0] }
	for parentCode in parentCodes:
		for score in scores[1:-1]:
			if parentCode in score[0]:
				report[parentCode]['score'] = [float(x) + float(y) for x, y in zip(report[parentCode]['score'], score[1:])]
				report[parentCode]['count'] += 1
		if report[parentCode]['count'] > 0:
			report[parentCode]['score'] = [ i / report[parentCode]['count'] for i in report[parentCode]['score'] ]
	report[scores[-1][0]] = {'count': 1, 'score': scores[-1][1:]}
	return report

#Function to save metrics into a csv file 
def putScoreCard(report, coderName):
	scoreCard = [['Code','Precision','Recall','F1','Support']]
	for code in sorted(report):
		scoreCard.append( [code] + report[code]['score'] )
	with open(coderName + '.csv', 'wb') as fp:
		wr = csv.writer(fp)
		wr.writerows(scoreCard)

#Function to check if the args supplied are valid files
def is_valid_file(parser, arg):
	if not os.path.exists(arg):
		parser.error("The file %s does not exist!" % arg)
	else:
		return arg

def main():
	parser = ArgumentParser(description="Calculate F1 Agreement Values Between Gold & Coder Sets")
	parser.add_argument("-gold", 
			    dest="gold", 
			    required=True,
			    help="Input a gold excel file", 
			    metavar="FILE",
	 		    type=lambda x: is_valid_file(parser, x))
	parser.add_argument("-coder",
			     dest="coder",
			     required=True,
			     help="Input a coder excel file",
			     metavar="FILE",																	        type=lambda x: is_valid_file(parser, x))
	args  = parser.parse_args()
        gold  = parseCSV(args.gold)
	coder = parseCSV(args.coder)
        report = getClassificationReport(gold, coder)
	putScoreCard(report, os.path.basename(args.coder).split('.')[0])


if __name__ == "__main__":
	main()
