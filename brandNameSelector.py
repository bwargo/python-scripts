# -*- coding: utf-8 -*-
import json
from pyexcel_xls import get_data
import csv
import itertools as it

data = get_data("brandnameAttributesForDataInput.xls")

rows = data["Sheet1"]
#print(json.dumps(rows))

transposedRows = [[row[i] for row in rows] for i in range(len(rows[0]))]

attributes = transposedRows.pop(0)

products = []

for x in range(len(transposedRows)):
	product = dict(zip(attributes,transposedRows[x]))
	products.append(product)

#print products
survey = {'Carpet':[1,3,4,5],'Hard Floor':[1,3,4,5],'Home Size':[1,3,4,5],'Powered Lift Away':[1,5],'Pets':[1,3,4,5],'Allergy & Asthma':[1,2,4,5],'Above Floor':[1,2,3,4,5]}


userSurvey = {'Carpet':1, 'Hard Floor':1,'Home Size':3,'Powered Lift Away':3,'Pets':1,'Allergy & Asthma':5,'Above Floor':2}

roundedUserSurvey = dict()

for key in userSurvey.iterkeys():
	roundedUserSurvey[key] = min(survey[key], key=lambda x:abs(x-userSurvey[key]))

#print roundedUserSurvey

def dict_compare(d1, d2):
    d1_keys = set(d1.keys())
    d2_keys = set(d2.keys())
    intersect_keys = d1_keys.intersection(d2_keys)
    added = d1_keys - d2_keys
    removed = d2_keys - d1_keys
    modified = {o : (d1[o], d2[o]) for o in intersect_keys if d1[o] != d2[o]}
    same = set(o for o in intersect_keys if d1[o] == d2[o])
    return added, removed, modified, same


#productOne = {'SKU': 'NV400', 'Hard Floor': 3, 'NAME': 'brandname Rotator Professional ', 'Powered Lift Away': 2, 'Above Floor': 4, 'Allergy & Asthma': 4, 'Pets': 1, 'Home Size': 4, 'Carpet': 5}

recommendedProductsSeven = []
recommendedProductsSix = []
recommendedProductsFive = []
recommendedProductsFour = []
recommendedProductsThree = []
recommendedProductsTwo = []
recommendedProductsOne = []

for prod in products:	
	added, removed, modified, same = dict_compare(userSurvey, prod)
	if len(same) > 6:
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

if len(recommendedProductsSeven) > 1:
	print [userSurvey, "Perfect Match. All Attributes Match for these products: ", recommendedProductsSeven]
elif len(recommendedProductsSix) > 1:
	print [userSurvey, "Six Matching Attributes for these products: ", recommendedProductsSix]
elif len(recommendedProductsFive) > 1:
	print [userSurvey, "Five Matching Attributes for these products: ", recommendedProductsFive]
elif len(recommendedProductsFour) > 1:
	print [userSurvey, "Four Matching Attributes for these products: ", recommendedProductsFour]
elif len(recommendedProductsThree) > 1:
	print [userSurvey, "Three Matching Attributes for these products: ", recommendedProductsThree]
elif len(recommendedProductsTwo) > 1:
	print [userSurvey, "Two Matching Attributes for these products: ", recommendedProductsTwo]
elif len(recommendedProductsOne) > 1:
	print [userSurvey, "One Matching Attribute for these products: ", recommendedProductsOne]








