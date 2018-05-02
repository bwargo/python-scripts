# -*- coding: utf-8 -*-
import json
from pyexcel_xls import get_data
import csv
import itertools as it

data = get_data("brandnameAttributesForDataInput922.xls")
storeData = get_data("brandnameStoresForDataInput922.xls")

rows = data["Sheet1"]
storeRows = storeData["Sheet1"]
fixedStoreRows = []
store = "Walmart.com"

#have to do this because rows may have different lengths
for storeRow in storeRows:
	diff_len = len(storeRows[0]) - len(storeRow)
	if diff_len < 0:
		raise AttributeError('Length of row is too long')
	fixedStoreRows.extend([storeRow + [''] * diff_len])

transposedRows = [[row[i] for row in rows] for i in range(len(rows[0]))]
transposedStoreRows = [[storeRow[j] for storeRow in fixedStoreRows] for j in range(len(fixedStoreRows[0]))]

attributes = transposedRows.pop(0)
storeAttributes = transposedStoreRows.pop(0)

products = []
storeProducts = []
storeSkus = []

for x in range(len(transposedStoreRows)):
	product = dict(zip(storeAttributes,transposedStoreRows[x]))
	#remove products that do not exist in store
	if product[store] != '':
		storeProducts.append(product)
		storeSkus.append(product['SKU'])
for x in range(len(transposedRows)):
	product = dict(zip(attributes,transposedRows[x]))
	#remove products that do not exist in store
	if product['SKU'] in storeSkus:
		products.append(product)

survey = {'Carpet':[1,3,4,5],'Hard Floor':[1,3,4,5],'Home Size':[1,3,4,5],'Powered Lift Away':[1,5],'Pets':[1,3,4,5],'Allergy & Asthma':[1,2,4,5],'Above Floor':[1,2,3,4,5],'Weight':[1,2,3,4,5]}


userSurvey = {'Carpet':1, 'Hard Floor':1,'Home Size':3,'Powered Lift Away':3,'Pets':1,'Allergy & Asthma':5,'Above Floor':2}

#roundedUserSurvey = dict()

#for key in userSurvey.iterkeys():
#	roundedUserSurvey[key] = min(survey[key], key=lambda x:abs(x-userSurvey[key]))

#print roundedUserSurvey

# calculate every possible combination of a survey
varNames = sorted(survey)
combinations = [dict(zip(varNames, metric)) for metric in it.product(*(survey[varName] for varName in varNames))]

#open up csv to write results to
b = open('recommendations922-withWeight-filteredTo5.csv', 'w')
combos = csv.writer(b)
combos.writerow(['User Selection','How Many Attributes Matched','Number of Original Products', 'Recommended Products (Max 5)'])

def dict_compare(d1, d2):
    d1_keys = set(d1.keys())
    d2_keys = set(d2.keys())
    intersect_keys = d1_keys.intersection(d2_keys)
    added = d1_keys - d2_keys
    removed = d2_keys - d1_keys
    modified = {o : (d1[o], d2[o]) for o in intersect_keys if d1[o] != d2[o]}
    same = set(o for o in intersect_keys if d1[o] == d2[o])
    return added, removed, modified, same

for combination in combinations:

	#productOne = {'SKU': 'NV400', 'Hard Floor': 3, 'NAME': 'brandname Rotator Professional ', 'Powered Lift Away': 2, 'Above Floor': 4, 'Allergy & Asthma': 4, 'Pets': 1, 'Home Size': 4, 'Carpet': 5}

	recommendedProductsEight = []
	recommendedProductsSeven = []
	recommendedProductsSix = []
	recommendedProductsFive = []
	recommendedProductsFour = []
	recommendedProductsThree = []
	recommendedProductsTwo = []
	recommendedProductsOne = []

	for prod in products:	
		added, removed, modified, same = dict_compare(combination, prod)
		if len(same) > 7:
			recommendedProductsEight.append([prod['SKU'],prod['NAME']])
		elif len(same) > 6:
			recommendedProductsSeven.append([prod['SKU'],prod['NAME']])
		elif len(same) > 5:
			recommendedProductsSix.append([prod['SKU'],prod['NAME']])
		elif len(same) > 4:
			recommendedProductsFive.append([prod['SKU'],prod['NAME']])
		elif len(same) > 3:
			recommendedProductsFour.append([prod['SKU'],prod['NAME']])
		elif len(same) > 2:
			recommendedProductsThree.append([prod['SKU'],prod['NAME']])
		elif len(same) > 1:
			recommendedProductsTwo.append([prod['SKU'],prod['NAME']])
		elif len(same) > 0:
			recommendedProductsOne.append([prod['SKU'],prod['NAME']])

	if len(recommendedProductsEight) > 0:
		combos.writerow([combination, "Perfect Match. All Attributes Match for these products: ", len(recommendedProductsEight), recommendedProductsEight[:5]])	
	elif len(recommendedProductsSeven) > 0:
		combos.writerow([combination, "Seven Matching Attributes for these products: ", len(recommendedProductsSeven), recommendedProductsSeven[:5]])
	elif len(recommendedProductsSix) > 0:
		combos.writerow([combination, "Six Matching Attributes for these products: ", len(recommendedProductsSix), recommendedProductsSix[:5]])
	elif len(recommendedProductsFive) > 0:
		combos.writerow([combination, "Five Matching Attributes for these products: ", len(recommendedProductsFive), recommendedProductsFive[:5]])
	elif len(recommendedProductsFour) > 0:
		combos.writerow([combination, "Four Matching Attributes for these products: ", len(recommendedProductsFour), recommendedProductsFour[:5]])
	elif len(recommendedProductsThree) > 0:
		combos.writerow([combination, "Three Matching Attributes for these products: ", len(recommendedProductsThree), recommendedProductsThree[:5]])
	elif len(recommendedProductsTwo) > 0:
		combos.writerow([combination, "Two Matching Attributes for these products: ", len(recommendedProductsTwo), recommendedProductsTwo[:5]])
	elif len(recommendedProductsOne) > 0:
		combos.writerow([combination, "One Matching Attribute for these products: ", len(recommendedProductsOne), recommendedProductsOne[:5]])

b.close()

#def findStores(sku):







