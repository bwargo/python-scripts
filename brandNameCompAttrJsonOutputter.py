# -*- coding: utf-8 -*-
import json
from pyexcel_xls import get_data
import csv
import itertools as it
import string

data = get_data("ProductSelectorFeatures_060817.xls")

rows = data["brandname Comparison Attributes"]
comparisonRows = data["Feature Bullets"]

#have to do this because rows may have different lengths
fixedStoreRows = []
for storeRow in rows:
	diff_len = len(rows[1]) - len(storeRow)
	if diff_len < 0:
		raise AttributeError('Length of row is too long')
	fixedStoreRows.extend([storeRow + [''] * diff_len])

categories = fixedStoreRows.pop(0) #removing first row of category names for now
tempCategory = ''
fixedCategories = []
#adding category name to each row
for category in categories: 
	if category != '':
		tempCategory = category
	else:
		category = tempCategory
	fixedCategories.append(category.lower())
fixedCategories[0] = 'franchise'

attributeRows = fixedStoreRows[0:10]
attributeRows.append(fixedCategories)
#print attributeRows
skuRows = fixedStoreRows[:1]+fixedStoreRows[11:]
print skuRows

#this groups together all attributes per product
#print fixedStoreRows
transposedProductAttributes = [[storeRow[i] for storeRow in attributeRows] for i in range(len(attributeRows[0]))]
#print '-------------------------'
#print transposedProductAttributes

transposedStoreSkus = [[storeRow[j] for storeRow in skuRows] for j in range(len(skuRows[0]))]
#print transposedStoreSkus

attributeNames = transposedProductAttributes.pop(0)
lcAttributeNames = []
for attributeName in attributeNames:
	lcAttributeNames.append(attributeName.lower())

#lets swap the names for normalized, lowercase names
stores = {"SKU":"base_sku","Amazon.com":"amazon_website","Best Buy (tablets in-store)":"bestbuy_instore","Bon-Ton":"bonton_instore","Kohl's (in-store)":"kohls_instore","Kohls.com":"kohls_website","Walmart.com":"walmart_website","brandname.com":"brandname_website"}
lcStoreNames = []
storeNames = transposedStoreSkus.pop(0)
for storeName in storeNames:
	lcStoreNames.append(stores[storeName])
#print attributeNames
products = {}

for x in range(len(transposedProductAttributes)):
	#label each attribute
	productWithAttributeName = dict(zip(lcAttributeNames,transposedProductAttributes[x]))
	#print productWithAttributeName	
	products[productWithAttributeName['sku']] = productWithAttributeName
#print products

for x in range(len(transposedStoreSkus)):
	#label each attribute
	productWithStoreName = dict(zip(lcStoreNames,transposedStoreSkus[x]))
	#print productWithStoreName
	#print productWithAttributeName	
	#lower-case all attribute names
	products[productWithStoreName['base_sku']]['stores'] = dict((k.lower(), v) for k,v in productWithStoreName.iteritems())

#now lets create a product entry for each store sku
storeProducts = {}
for key, value in products.iteritems():
	storeValues = value["stores"]
	for storeKey, storeValue in storeValues.iteritems():
		storeValueList = storeValue.split(",")
		#commenting this out, we are only keeping the 1st sku if there are multiple skus 1-31-17 BW
		#for val in storeValueList:
		#	fixedVal = val.strip()
		fixedVal = storeValueList[0].strip()
		if fixedVal not in storeProducts.keys() and fixedVal != '':
			storeProducts[fixedVal] = value

# OK, now lets do the third tab 'Feature Bullets'!
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

#remove doubled up columns since some columns are "hidden" in the spreadsheet
finalProductFeatures = {}
for row in transposedProductFeatures:
	if row[0] != "" and row[0] not in finalProductFeatures:
		finalProductFeatures[row[0]] = row[1:]

#alright now lets put it all together! 
for featureKey, featureValue in finalProductFeatures.iteritems():
	if featureKey not in storeProducts.keys():
		print featureKey
	else:
		#remove bullets
		fixedFeatureValue = []
		for feature in featureValue:
			#print feature
			if(len(feature) > 0):
				fixedFeatureValue.append(string.replace(feature, u"\u2022", ''))
		storeProducts[featureKey]["features"] = fixedFeatureValue

#hopefully a temp fix for missing features for NV401 and NV801, using NV400 and NV800's product features for them
for splatKey, splatValue in storeProducts.iteritems():
	if splatKey == "NV401":
		storeProducts[splatKey]["features"] = storeProducts["NV400"]["features"]
	if splatKey == "NV801":
		storeProducts[splatKey]["features"] = storeProducts["NV800"]["features"]

#last safety check if "features" are missing 5-10-2017
for splatKey, splatValue in storeProducts.iteritems():
	if "features" not in splatValue.keys():
		storeProducts[splatKey]["features"] = []

#open up file to write results to
with open("productAttributes.js", "w") as outfile:
	outfile.write("var productAttributes=")
	json.dump(storeProducts, outfile, indent=1)
	outfile.write(";")
