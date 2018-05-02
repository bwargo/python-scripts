# -*- coding: utf-8 -*-
import json
from pyexcel_xls import get_data
import csv
import itertools as it

data = get_data("ProductSelectorFeatures_060817.xls")

rows = data["brandname Selector Questions"]
rows.pop(0) #removing first row of category names for now
questionRows = rows[0:11]
storeRows = rows[12:19]
parentSkus = rows[0]
#print parentSkus

#default store , the store must match the name exactly in the spreadsheet. ex: brandname.com, Amazon.com, 
#Best Buy, Bon Ton, Kohls.com, Walmart.com
stores = {"amazon_website": "Amazon.com","bestbuy_instore":"Best Buy (in-store)","bonton_instore":"Bon-Ton","kohls_instore":"Kohl's (in-store)","kohls_website":"Kohls.com","walmart_website":"Walmart.com","brandname_website":"brandname.com"}
for storeKey, storeValue in stores.iteritems():
	fileName = "surveykeys-"+storeKey+".js"
	#have to do this because rows may have different lengths
	fixedStoreRows = []
	for storeRow in storeRows:
		#print storeRow
		diff_len = len(rows[0]) - len(storeRow)
		if diff_len < 0:
			raise AttributeError('Length of row is too long')
		fixedStoreRows.extend([storeRow + [''] * diff_len])

	#this groups together all question's ratings per product
	#print questionRows
	transposedProductRatings = [[row[i] for row in questionRows] for i in range(len(questionRows[0]))]
	#print '-------------------------'
	#print transposedProductRatings

	#this groups together all store skus per product
	#print fixedStoreRows
	transposedStoreSkus = [[storeRow[j] for storeRow in fixedStoreRows] for j in range(len(fixedStoreRows[0]))]
	#print '-------------------------'
	#print transposedStoreSkus

	questionNames = transposedProductRatings.pop(0)
	#print questionNames
	storeNames = transposedStoreSkus.pop(0)
	#print storeNames
	#print transposedStoreSkus

	products = []
	#storeProducts = []
	storeSkus = {}

	for x in range(len(transposedStoreSkus)):
		#label each sku with a store name
		productWithStoreName = dict(zip(storeNames,transposedStoreSkus[x]))
		#print productWithStoreName
		#remove products that do not exist in store
		print transposedStoreSkus[x]
		print storeValue	
		if productWithStoreName[storeValue] != '':
			#storeProducts.append(product)
			storeSkus[productWithStoreName[storeValue]] = productWithStoreName[storeValue]
	#print storeSkus
	for x in range(len(transposedProductRatings)):
		#label each Product's Rating with a Question Name
		product = dict(zip(questionNames,transposedProductRatings[x]))
		#print product
		if product['SKU'] in storeSkus:
			#print product['SKU']
			products.append(product)
	#print len(products)
	#all Questions and their possible Ratings
	survey = {'Carpet':[1,2,3,4,5],'Hard Floor':[1,2,3,4,5],'Home Size':[1,2,3,4,5],'Powered Lift Away':[1,5],'Pets':[1,3,4,5],'Allergy & Asthma':[1,2,4,5],'Above Floor':[1,2,3,4,5], 'Weight':[1,2,3,4,5]}

	#next few lines using 'userSurvey' are for testing
	#userSurvey = {'Carpet':1, 'Hard Floor':1,'Home Size':3,'Powered Lift Away':3,'Pets':1,'Allergy & Asthma':5,'Above Floor':2}
	#roundedUserSurvey = dict()
	#for key in userSurvey.iterkeys():
	#	roundedUserSurvey[key] = min(survey[key], key=lambda x:abs(x-userSurvey[key]))
	#print roundedUserSurvey

	# calculate every possible combination of a survey
	varNames = sorted(survey)
	combinations = [dict(zip(varNames, metric)) for metric in it.product(*(survey[varName] for varName in varNames))]

	def dict_compare(d1, d2):
	    d1_keys = set(d1.keys())
	    d2_keys = set(d2.keys())
	    intersect_keys = d1_keys.intersection(d2_keys)
	    added = d1_keys - d2_keys
	    removed = d2_keys - d1_keys
	    modified = {o : (d1[o], d2[o]) for o in intersect_keys if d1[o] != d2[o]}
	    same = set(o for o in intersect_keys if d1[o] == d2[o])
	    return added, removed, modified, same

	#open up file to write results to
	with open(fileName, "w") as outfile:
		#b = open('recommendations922-withWeight-filteredTo5.csv', 'w')
		#combos = csv.writer(b)
		#combos.writerow(['User Selection','How Many Attributes Matched','Number of Original Products', 'Recommended Products (Max 5)'])
		combosToPrint = {}
		outfile.write("var surveykeys={")
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
					recommendedProductsEight.append(prod['SKU'])
				elif len(same) > 6:
					recommendedProductsSeven.append(prod['SKU'])
				elif len(same) > 5:
					recommendedProductsSix.append(prod['SKU'])
				elif len(same) > 4:
					recommendedProductsFive.append(prod['SKU'])
				elif len(same) > 3:
					recommendedProductsFour.append(prod['SKU'])
				elif len(same) > 2:
					recommendedProductsThree.append(prod['SKU'])
				elif len(same) > 1:
					recommendedProductsTwo.append(prod['SKU'])
				elif len(same) > 0:
					recommendedProductsOne.append(prod['SKU'])

			combo = ''
			for x in combination:
				combo = combo+str(combination[x])
			if len(recommendedProductsEight) > 0:
				combosToPrint[combo] = recommendedProductsEight
				#json.dump({combo:recommendedProductsEight}, outfile)
				#combos.writerow([combination, "Perfect Match. All Attributes Match for these products: ", len(recommendedProductsEight), recommendedProductsEight[:5]])	
			elif len(recommendedProductsSeven) > 0:
				combosToPrint[combo] = recommendedProductsSeven
				#json.dump({combo:recommendedProductsSeven}, outfile)
				#combos.writerow([combination, "Seven Matching Attributes for these products: ", len(recommendedProductsSeven), recommendedProductsSeven[:5]])
			elif len(recommendedProductsSix) > 0:
				combosToPrint[combo] = recommendedProductsSix
				#json.dump({combo:recommendedProductsSix}, outfile)
				#combos.writerow([combination, "Six Matching Attributes for these products: ", len(recommendedProductsSix), recommendedProductsSix[:5]])
			elif len(recommendedProductsFive) > 0:
				combosToPrint[combo] = recommendedProductsFive
				#json.dump({combo:recommendedProductsFive}, outfile)
				#combos.writerow([combination, "Five Matching Attributes for these products: ", len(recommendedProductsFive), recommendedProductsFive[:5]])
			elif len(recommendedProductsFour) > 0:
				combosToPrint[combo] = recommendedProductsFour
				#json.dump({combo:recommendedProductsFour}, outfile)
				#combos.writerow([combination, "Four Matching Attributes for these products: ", len(recommendedProductsFour), recommendedProductsFour[:5]])
			elif len(recommendedProductsThree) > 0:
				combosToPrint[combo] = recommendedProductsThree
				#json.dump({combo:recommendedProductsThree}, outfile)
				#combos.writerow([combination, "Three Matching Attributes for these products: ", len(recommendedProductsThree), recommendedProductsThree[:5]])
			elif len(recommendedProductsTwo) > 0:
				combosToPrint[combo] = recommendedProductsTwo
				#json.dump({combo:recommendedProductsTwo}, outfile)
				#combos.writerow([combination, "Two Matching Attributes for these products: ", len(recommendedProductsTwo), recommendedProductsTwo[:5]])
			elif len(recommendedProductsOne) > 0:
				combosToPrint[combo] = recommendedProductsOne
				#json.dump({combo:recommendedProductsOne}, outfile)
			    #json.dump({'numbers':combination, 'products':recommendedProductsOne}, outfile, indent=4)
				#combos.writerow([combination, "One Matching Attribute for these products: ", len(recommendedProductsOne), recommendedProductsOne[:5]])
		
		last = combosToPrint.items()[-1]
		print last
		for k,v in combosToPrint.items():
			json.dump(int(k), outfile)
			outfile.write(":")
			json.dump(v,outfile)
			if(k!=last[0]):
				outfile.write(",")
		outfile.write("};")	
#b.close()

#def findStores(sku):







