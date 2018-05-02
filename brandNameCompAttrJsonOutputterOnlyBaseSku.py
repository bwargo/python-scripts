# -*- coding: utf-8 -*-
import json
from pyexcel_xls import get_data
import csv
import itertools as it

data = get_data("ProductSelectorFeatures_RJM_010317.xls")

rows = data["brandname Comparison Attributes"]
rows.pop(0) #removing first row of category names for now

#have to do this because rows may have different lengths
fixedStoreRows = []
for storeRow in rows:
	#print storeRow
	diff_len = len(rows[0]) - len(storeRow)
	if diff_len < 0:
		raise AttributeError('Length of row is too long')
	fixedStoreRows.extend([storeRow + [''] * diff_len])

#this groups together all attributes per product
#print fixedStoreRows
transposedProductAttributes = [[storeRow[j] for storeRow in fixedStoreRows] for j in range(len(fixedStoreRows[0]))]
#print '-------------------------'
#print transposedProductAttributes

attributeNames = transposedProductAttributes.pop(0)
#print attributeNames
products = {}

for x in range(len(transposedProductAttributes)):
	#label each attribute
	productWithAttributeName = dict(zip(attributeNames,transposedProductAttributes[x]))
	#print productWithAttributeName	
	products[productWithAttributeName['SKU']] = productWithAttributeName
#print products

#open up file to write results to
with open("productAttributes.json", "w") as outfile:
	json.dump(products, outfile, indent=1)
