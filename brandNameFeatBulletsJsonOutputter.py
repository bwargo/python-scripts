# -*- coding: utf-8 -*-
import json
from pyexcel_xls import get_data
import csv
import itertools as it

data = get_data("ProductSelectorFeatures_RJM_010317.xls")

comparisonRows = data["Feature Bullets"]

#have to do this because rows may have different lengths
fixedComparisonRows = []
for comparisonRow in comparisonRows:
	#print comparisonRow
	diff_len = len(comparisonRows[0]) - len(comparisonRow)
	if diff_len < 0:
		raise AttributeError('Length of row is too long')
	fixedComparisonRows.extend([comparisonRow + [''] * diff_len])

#print fixedComparisonRows
transposedProductFeatures = [[row[j] for row in fixedComparisonRows] for j in range(len(fixedComparisonRows[0]))]
#print transposedProductFeatures

headings = transposedProductFeatures.pop(0)


#remove doubled up columns
finalProductFeatures = {}

for row in transposedProductFeatures:
	if row[0] != "" and row[0] not in finalProductFeatures:
		finalProductFeatures[row[0]] = row[1:]
fileName = "productFeatures.json"
with open(fileName, "w") as outfile:
	json.dump(finalProductFeatures, outfile, indent=1)