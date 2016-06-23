#!/usr/bin/python

# Name: genExcelByLocalityIds.py
# Author: [Govindrao Kulkarni]
# Description: Generate 'Keywords' and 'Ads' Excel Sheets using locality ids
# Prerequisite: Install xlrd and xlsxwriter libraires
#               Edit the config, inputFilePath: excel sheet with list of localityIds
#                                performCountCheck: if True, will include adGroup row if listing count is > 5

import xlrd
import xlsxwriter
import requests
import json

config = {
    'inputFilePath': "/home/govind/work/myExps/pythonScripts/localityIds.xlsx",
    'performCountCheck': False
}

keep_all_rows = config['performCountCheck']
totalCountMap = {}
bhk1CountMap = {}
bhk2CountMap = {}
bhk3CountMap = {}
bhk4CountMap = {}

exceptionIdList = []

s1_rowNum = 0
s2_rowNum = 0

# Create a output workbook and add a worksheet.
workbook = xlsxwriter.Workbook('marketing.xlsx')
worksheet1 = workbook.add_worksheet('keywords')
worksheet2 = workbook.add_worksheet('ads')

def scriptInit():
    locId_file = config['inputFilePath']
    wb1 = xlrd.open_workbook(locId_file)
    sheet1 = wb1.sheet_by_index(0)
    if keep_all_rows == True:
        fetchLocalityCountData()
    for i in range(1, sheet1.nrows):
        getLocalityLabels(int(sheet1.cell_value(i,0)))
    print exceptionIdList
    workbook.close()

def  apiCaller(api):
    print 'GET call: ', api
    res = requests.get(api)
    if res.status_code == requests.codes.ok:
        return json.loads(res.content)
    else:
        print 'ERROR in GET call: ', api
        # If response code is not ok (200), print the resulting http error code with description
        res.raise_for_status()
        return None

def fetchLocalityCountData():
    apiList = {
        'property_api' : 'http://proptiger.com/app/v1/listing?selector={"filters":{"and":[{"equal":{"listingCategory":["Primary","Resale"]}}]},"paging":{"start":0,"rows":0}}&facets=localityId&sourceDomain=Makaan',
        # 'apartment_api' : 'http://www.proptiger.com/app/v1/listing?selector={"filters":{"and":[{"equal":{"listingCategory":["Primary","Resale"]}},{"range":{"price":{"from":"0","to":"5000000"}}}]},"paging":{"start":0,"rows":0}}&facets=localityId&sourceDomain=Makaan',
        # 'homes_api' : 'http://www.proptiger.com/app/v1/listing?selector={"filters":{"and":[{"equal":{"listingCategory":["Primary","Resale"]}},{"range":{"price":{"from":"0","to":"5000000"}}}]},"paging":{"start":0,"rows":0}}&facets=localityId&sourceDomain=Makaan',
        # 'houses_api' : 'http://www.proptiger.com/app/v1/listing?selector={"filters":{"and":[{"equal":{"listingCategory":["Primary","Resale"]}},{"range":{"price":{"from":"0","to":"5000000"}}}]},"paging":{"start":0,"rows":0}}&facets=localityId&sourceDomain=Makaan',
        # 'flats_api' : 'http://www.proptiger.com/app/v1/listing?selector={"filters":{"and":[{"equal":{"listingCategory":["Primary","Resale"]}},{"range":{"price":{"from":"0","to":"5000000"}}}]},"paging":{"start":0,"rows":0}}&facets=localityId&sourceDomain=Makaan',
        'bhk2_api' : 'http://www.proptiger.com/app/v1/listing?selector={"filters":{"and":[{"equal":{"listingCategory":["Primary","Resale"]}},{"equal":{"bedrooms":["2"]}},{"range":{"price":{"from":"0","to":"5000000"}}}]},"paging":{"start":0,"rows":0}}&facets=localityId&sourceDomain=Makaan',
        'bhk3_api' : 'http://www.proptiger.com/app/v1/listing?selector={"filters":{"and":[{"equal":{"listingCategory":["Primary","Resale"]}},{"equal":{"bedrooms":["3"]}},{"range":{"price":{"from":"0","to":"5000000"}}}]},"paging":{"start":0,"rows":0}}&facets=localityId&sourceDomain=Makaan',
        'bhk4_api' : 'http://www.proptiger.com/app/v1/listing?selector={"filters":{"and":[{"equal":{"listingCategory":["Primary","Resale"]}},{"equal":{"bedrooms":["4"]}},{"range":{"price":{"from":"0","to":"5000000"}}}]},"paging":{"start":0,"rows":0}}&facets=localityId&sourceDomain=Makaan',
        'bhk1_api' : 'http://www.proptiger.com/app/v1/listing?selector={"filters":{"and":[{"equal":{"listingCategory":["Primary","Resale"]}},{"equal":{"bedrooms":["1"]}},{"range":{"price":{"from":"0","to":"5000000"}}}]},"paging":{"start":0,"rows":0}}&facets=localityId&sourceDomain=Makaan'
        # 'budget_property_api' : 'https://www.proptiger.com/app/v1/listing?selector={"filters":{"and":[{"equal":{"listingCategory":["Primary","Resale"]}},{"range":{"price":{"from":"0","to":"5000000"}}}]},"paging":{"start":0,"rows":0}}&facets=localityId&sourceDomain=Makaan',
        # 'affordable_property_api' : 'https://www.proptiger.com/app/v1/listing?selector={"filters":{"and":[{"equal":{"listingCategory":["Primary","Resale"]}},{"range":{"price":{"from":"0","to":"5000000"}}}]},"paging":{"start":0,"rows":0}}&facets=localityId&sourceDomain=Makaan',
        # 'builder_floor_api' : 'https://www.proptiger.com/app/v1/listing?selector={"filters":{"and":[{"equal":{"listingCategory":["Primary","Resale"]}},{"range":{"price":{"from":"0","to":"5000000"}}}]},"paging":{"start":0,"rows":0}}&facets=localityId&sourceDomain=Makaan'
    }

    for key, val in enumerate(apiList):
        apiData = apiCaller(apiList[val])
        if apiData is not None:
            createMap(apiData, val)

def createMap(apiData, key):
    if apiData.has_key('data') and apiData['data'].has_key('facets') and apiData['data']['facets'].has_key('localityId'):
        locArr = apiData['data']['facets']['localityId']
        if key == 'property_api' or key == 'apartment_api' or key == 'homes_api' or key == 'houses_api' or key == 'flats_api' or key == 'budget_property_api' or key == 'affordable_property_api' or key == 'builder_floor_api':
            global totalCountMap
            populateMapObj(locArr, totalCountMap)
        elif key == 'bhk1_api':
            global bhk1CountMap
            populateMapObj(locArr, bhk1CountMap)
        elif key == 'bhk2_api':
            global bhk2CountMap
            populateMapObj(locArr, bhk2CountMap)
        elif key == 'bhk3_api':
            global bhk3CountMap
            populateMapObj(locArr, bhk3CountMap)
        elif key == 'bhk4_api':
            global bhk4CountMap
            populateMapObj(locArr, bhk4CountMap)

def populateMapObj(locArr, countIdMap):
    for countObj in locArr:
        for locId in countObj:
            countIdMap[locId] = countObj[locId]

def getLocalityLabels(locId):
    print locId
    url = 'https://www.proptiger.com/app/v2/locality/' + `locId` + '?selector={"fields":["localityId","suburb","city","label"]}&sourceDomain=Makaan'
    apiData = apiCaller(url)
    if apiData is not None:
        if apiData.has_key('data') and apiData['data'] is not None and apiData['data'].has_key('label'):
            localityLabel = apiData['data']['label']
        else:
            return exceptionIdList.append(locId)
        if apiData.has_key('data') and apiData['data'] is not None and apiData['data'].has_key('suburb') and apiData['data']['suburb'].has_key('city') and apiData['data']['suburb']['city'].has_key('label'):
            cityLabel = apiData['data']['suburb']['city']['label']
        else:
            return exceptionIdList.append(locId)
        generateCurrLocalityContent(cityLabel, localityLabel, locId)

def getToBeAddedData(locId, locLabel):
    return {
            locLabel + ' ' + 'Property': keep_all_rows or (totalCountMap.has_key(`locId`) and totalCountMap[`locId`] > 5),
            locLabel + ' ' + 'Apartments': keep_all_rows or (totalCountMap.has_key(`locId`) and totalCountMap[`locId`] > 5),
            locLabel + ' ' + 'Homes': keep_all_rows or (totalCountMap.has_key(`locId`) and totalCountMap[`locId`] > 5),
            locLabel + ' ' + 'House': keep_all_rows or (totalCountMap.has_key(`locId`) and totalCountMap[`locId`] > 5),
            locLabel + ' ' + 'flats': keep_all_rows or (totalCountMap.has_key(`locId`) and totalCountMap[`locId`] > 5),
            locLabel + ' ' + 'Builder Floors': keep_all_rows or (totalCountMap.has_key(`locId`) and totalCountMap[`locId`] > 5),
            '1BHK' + ' ' + locLabel: keep_all_rows or (bhk1CountMap.has_key(`locId`) and bhk1CountMap[`locId`] > 5),
            '2BHK' + ' ' + locLabel: keep_all_rows or (bhk2CountMap.has_key(`locId`) and bhk2CountMap[`locId`] > 5),
            '3BHK' + ' ' + locLabel: keep_all_rows or (bhk3CountMap.has_key(`locId`) and bhk3CountMap[`locId`] > 5),
            '4BHK' + ' ' + locLabel: keep_all_rows or (bhk4CountMap.has_key(`locId`) and bhk4CountMap[`locId`] > 5),
            'Budget Proeprty' + ' ' + locLabel: keep_all_rows or (totalCountMap.has_key(`locId`) and totalCountMap[`locId`] > 5),
            'Affordable Property' + ' ' + locLabel: keep_all_rows or (totalCountMap.has_key(`locId`) and totalCountMap[`locId`] > 5)
        }

def getLastColumnUrl(locId):
    lastColUrls = {}
    api1 = 'https://www.makaan.com/dawnstar/data/v2/fetch-urls?urlParam=[{"urlDomain":"locality","domainIds":[' + `locId` + '],"urlCategoryName":"MAKAAN_LOCALITY_BHK_PROPERTY_BUY"}]'
    apiData = apiCaller(api1)
    if apiData is not None:
        if apiData.has_key('data') and apiData['data'] is not None and apiData['data'].has_key('MAKAAN_LOCALITY_BHK_PROPERTY_BUY-' + `locId`):
            lastColUrls['propertyBhkUrl'] = 'https://www.makaan.com/' + apiData['data']['MAKAAN_LOCALITY_BHK_PROPERTY_BUY-' + `locId`]
        else:
            print 'error in fetching propertyBhkUrl for ', locId
    api2 = 'https://www.makaan.com/dawnstar/data/v2/fetch-urls?urlParam=[{"urlDomain":"locality","domainIds":[' + `locId` + '],"urlCategoryName":"MAKAAN_LOCALITY_LISTING_BUY"}]'
    apiData = apiCaller(api2)
    if apiData is not None:
        if apiData.has_key('data') and apiData['data'] is not None and apiData['data'].has_key('MAKAAN_LOCALITY_LISTING_BUY-' + `locId`):
            lastColUrls['propertyUrl'] = 'https://www.makaan.com/' + apiData['data']['MAKAAN_LOCALITY_LISTING_BUY-' + `locId`]
        else:
            print 'error in fetching propertyUrl for ', locId
    return lastColUrls


def populateAdsWorksheet(adgNum, adGroup, locLabel, campaign, locId, lcUrls):
    worksheet2.write(s2_rowNum, 0, campaign)
    if adgNum == 0:
        worksheet2.write(s2_rowNum, 1, adGroup[adgNum])
        worksheet2.write(s2_rowNum, 2, locLabel + ' Property')
        worksheet2.write(s2_rowNum, 3, 'Huge Options in Multiple Budgets.')
        worksheet2.write(s2_rowNum, 4, 'Best Price only on Makaan.com')
        worksheet2.write(s2_rowNum, 5, 'Makaan.com/' + locLabel.replace(" ", "") + '_Property')
        worksheet2.write(s2_rowNum, 6, lcUrls['propertyUrl'])
    elif adgNum == 1:
        worksheet2.write(s2_rowNum, 1, adGroup[adgNum])
        worksheet2.write(s2_rowNum, 2, locLabel + ' Apartment')
        worksheet2.write(s2_rowNum, 3, 'Huge Options in Multiple Budgets.')
        worksheet2.write(s2_rowNum, 4, 'Best Price only on Makaan.com')
        worksheet2.write(s2_rowNum, 5, 'Makaan.com/' + locLabel.replace(" ", "") + '_Apts')
        worksheet2.write(s2_rowNum, 6, lcUrls['propertyUrl'])
    elif adgNum == 2:
        worksheet2.write(s2_rowNum, 1, adGroup[adgNum])
        worksheet2.write(s2_rowNum, 2, locLabel + ' Home')
        worksheet2.write(s2_rowNum, 3, 'Huge Options in Multiple Budgets.')
        worksheet2.write(s2_rowNum, 4, 'Best Price only on Makaan.com')
        worksheet2.write(s2_rowNum, 5, 'Makaan.com/' + locLabel.replace(" ", "_") + '_Home')
        worksheet2.write(s2_rowNum, 6, lcUrls['propertyUrl'])
    elif adgNum == 3:
        worksheet2.write(s2_rowNum, 1, adGroup[adgNum])
        worksheet2.write(s2_rowNum, 2, locLabel + ' House')
        worksheet2.write(s2_rowNum, 3, '100% Real & Verified Properties.')
        worksheet2.write(s2_rowNum, 4, 'Find by Location & Budget Now!')
        worksheet2.write(s2_rowNum, 5, 'Makaan.com/' + locLabel.replace(" ", "_") + '_House')
        worksheet2.write(s2_rowNum, 6, lcUrls['propertyUrl'])
    elif adgNum == 4:
        worksheet2.write(s2_rowNum, 1, adGroup[adgNum])
        worksheet2.write(s2_rowNum, 2, 'Flats in ' + locLabel)
        worksheet2.write(s2_rowNum, 3, '100% Real & Verified Properties.')
        worksheet2.write(s2_rowNum, 4, 'Find by Location & Budget Now!')
        worksheet2.write(s2_rowNum, 5, 'Makaan.com/Flats_' + locLabel.replace(" ", "_"))
        worksheet2.write(s2_rowNum, 6, lcUrls['propertyUrl'])
    elif adgNum == 5:
        worksheet2.write(s2_rowNum, 1, adGroup[adgNum])
        worksheet2.write(s2_rowNum, 2, locLabel + ' Property')
        worksheet2.write(s2_rowNum, 3, 'Buy Builder Floor, Best Options.')
        worksheet2.write(s2_rowNum, 4, 'Best Price only on Makaan.com')
        worksheet2.write(s2_rowNum, 5, 'Makaan.com/' + locLabel.replace(" ", "") + '_Property')
        worksheet2.write(s2_rowNum, 6, lcUrls['propertyUrl'])
    elif adgNum == 6:
        worksheet2.write(s2_rowNum, 1, adGroup[adgNum])
        worksheet2.write(s2_rowNum, 2, locLabel + ' 1 BHK')
        worksheet2.write(s2_rowNum, 3, '100% Real & Verified Properties.')
        worksheet2.write(s2_rowNum, 4, 'Find by Location & Budget Now!')
        worksheet2.write(s2_rowNum, 5, 'Makaan.com/' + locLabel.replace(" ", "_") + '_1_BHK')
        worksheet2.write(s2_rowNum, 6, lcUrls['propertyBhkUrl'].replace('-bhk-','-1bhk-'))
    elif adgNum == 7:
        worksheet2.write(s2_rowNum, 1, adGroup[adgNum])
        worksheet2.write(s2_rowNum, 2, locLabel + ' 2 BHK')
        worksheet2.write(s2_rowNum, 3, 'Huge Options in Multiple Budgets.')
        worksheet2.write(s2_rowNum, 4, 'Best Price only on Makaan.com')
        worksheet2.write(s2_rowNum, 5, 'Makaan.com/' + locLabel.replace(" ", "_") + '_2_BHK')
        worksheet2.write(s2_rowNum, 6, lcUrls['propertyBhkUrl'].replace('-bhk-','-2bhk-'))
    elif adgNum == 8:
        worksheet2.write(s2_rowNum, 1, adGroup[adgNum])
        worksheet2.write(s2_rowNum, 2, locLabel + ' 2 BHK')
        worksheet2.write(s2_rowNum, 3, '100% Real & Verified Properties.')
        worksheet2.write(s2_rowNum, 4, 'Find by Location & Budget Now!')
        worksheet2.write(s2_rowNum, 5, 'Makaan.com/' + locLabel.replace(" ", "_") + '_3_BHK')
        worksheet2.write(s2_rowNum, 6, lcUrls['propertyBhkUrl'].replace('-bhk-','-3bhk-'))
    elif adgNum == 9:
        worksheet2.write(s2_rowNum, 1, adGroup[adgNum])
        worksheet2.write(s2_rowNum, 2, locLabel + ' 4 BHK')
        worksheet2.write(s2_rowNum, 3, 'Huge Options in Multiple Budgets.')
        worksheet2.write(s2_rowNum, 4, 'Best Price only on Makaan.com')
        worksheet2.write(s2_rowNum, 5, 'Makaan.com/' + locLabel.replace(" ", "_") + '_4_BHK')
        worksheet2.write(s2_rowNum, 6, lcUrls['propertyBhkUrl'].replace('-bhk-','-4bhk-'))
    elif adgNum == 10:
        worksheet2.write(s2_rowNum, 1, adGroup[adgNum])
        worksheet2.write(s2_rowNum, 2, locLabel + ' Property')
        worksheet2.write(s2_rowNum, 3, 'Property for Every Budget.')
        worksheet2.write(s2_rowNum, 4, 'Best Price only on Makaan.com')
        worksheet2.write(s2_rowNum, 5, 'Makaan.com/' + locLabel.replace(" ", "") + '_Property')
        worksheet2.write(s2_rowNum, 6, lcUrls['propertyUrl'])
    elif adgNum == 11:
        worksheet2.write(s2_rowNum, 1, adGroup[adgNum])
        worksheet2.write(s2_rowNum, 2, locLabel + ' Property')
        worksheet2.write(s2_rowNum, 3, 'Wide Range of Affordable Properties')
        worksheet2.write(s2_rowNum, 4, 'Find by Location & Budget Now!')
        worksheet2.write(s2_rowNum, 5, 'Makaan.com/' + locLabel.replace(" ", "") + '_Property')
        worksheet2.write(s2_rowNum, 6, lcUrls['propertyUrl'])


def generateCurrLocalityContent(cityLabel, locLabel, locId):
    global worksheet1
    global s1_rowNum
    global s2_rowNum

    campaign = cityLabel + '-' + locLabel
    toBeAdded = getToBeAddedData(locId, locLabel)
    print toBeAdded
    adGroup = [
        locLabel + ' ' + 'Property',
        locLabel + ' ' + 'Apartments',
        locLabel + ' ' + 'Homes',
        locLabel + ' ' + 'House',
        locLabel + ' ' + 'flats',
        locLabel + ' ' + 'Builder Floors',
        '1BHK' + ' ' + locLabel,
        '2BHK' + ' ' + locLabel,
        '3BHK' + ' ' + locLabel,
        '4BHK' + ' ' + locLabel,
        'Budget Proeprty' + ' ' + locLabel,
        'Affordable Property' + ' ' + locLabel
    ]

    keyWordsColumn = {
        adGroup[0]: [
            '+property in +' + locLabel.replace(' ', ' +'),
            '+properties in +' + locLabel.replace(' ', ' +'),
            '+' + locLabel.replace(' ', ' +') + ' +property',
            '+' + locLabel.replace(' ', ' +') + ' +properties',
            '+property for +sale in +' + locLabel.replace(' ', ' +'),
            '+properties for +sale in +' + locLabel.replace(' ', ' +'),
            '+buy +property  in +' + locLabel.replace(' ', ' +'),
            '+' + locLabel.replace(' ', ' +') + ' +property for +sale',
            '+' + locLabel.replace(' ', ' +') + ' +property'
        ],
        adGroup[1]: [
            '+apartment in +' + locLabel.replace(' ', ' +'),
            '+apartments in +' + locLabel.replace(' ', ' +'),
            '+' + locLabel.replace(' ', ' +') + ' +apartment',
            '+' + locLabel.replace(' ', ' +') + ' +apartments',
            '+apartment for +sale in +' + locLabel.replace(' ', ' +'),
            '+apartments for +sale in +' + locLabel.replace(' ', ' +'),
            '+buy +apartment  in +' + locLabel.replace(' ', ' +'),
            '+' + locLabel.replace(' ', ' +') + ' +apartment for +sale',
            '+' + locLabel.replace(' ', ' +') + ' +apartment'
        ],
        adGroup[2]: [
            '+home in +' + locLabel.replace(' ', ' +'),
            '+homes in +' + locLabel.replace(' ', ' +'),
            '+' + locLabel.replace(' ', ' +') + ' +home',
            '+' + locLabel.replace(' ', ' +') + ' +homes',
            '+home for +sale in +' + locLabel.replace(' ', ' +'),
            '+homes for +sale in +' + locLabel.replace(' ', ' +'),
            '+buy +home  in +' + locLabel.replace(' ', ' +'),
            '+' + locLabel.replace(' ', ' +') + ' +home for +sale',
            '+' + locLabel.replace(' ', ' +') + ' +home'
        ],
        adGroup[3]: [
            '+house in +' + locLabel.replace(' ', ' +'),
            '+houses in +' + locLabel.replace(' ', ' +'),
            '+' + locLabel.replace(' ', ' +') + ' +house',
            '+' + locLabel.replace(' ', ' +') + ' +houses',
            '+house for +sale in +' + locLabel.replace(' ', ' +'),
            '+houses for +sale in +' + locLabel.replace(' ', ' +'),
            '+buy +house  in +' + locLabel.replace(' ', ' +'),
            '+' + locLabel.replace(' ', ' +') + ' +house for +sale',
            '+' + locLabel.replace(' ', ' +') + ' +house'
        ],
        adGroup[4]: [
            '+flat in +' + locLabel.replace(' ', ' +'),
            '+flats in +' + locLabel.replace(' ', ' +'),
            '+' + locLabel.replace(' ', ' +') + ' +flat',
            '+' + locLabel.replace(' ', ' +') + ' +flats',
            '+flat for +sale in +' + locLabel.replace(' ', ' +'),
            '+flats for +sale in +' + locLabel.replace(' ', ' +'),
            '+buy +flat  in +' + locLabel.replace(' ', ' +'),
            '+' + locLabel.replace(' ', ' +') + ' +flat for +sale',
            '+' + locLabel.replace(' ', ' +') + ' +flat'
        ],
        adGroup[5]: [
            '+builder +floor in +' + locLabel.replace(' ', ' +'),
            '+builder +floors in +' + locLabel.replace(' ', ' +'),
            '+' + locLabel.replace(' ', ' +') + ' +builder +floor',
            '+' + locLabel.replace(' ', ' +') + ' +builder +floors',
            '+builder +floor for +sale in +' + locLabel.replace(' ', ' +'),
            '+builder +floors for +sale in +' + locLabel.replace(' ', ' +'),
            '+buy +builder +floor  in +' + locLabel.replace(' ', ' +'),
            '+' + locLabel.replace(' ', ' +') + ' +builder +floor for +sale',
            '+' + locLabel.replace(' ', ' +') + ' +builder +floor'
        ],
        adGroup[6]: [
            '+1 +BHK +' + locLabel.replace(' ', ' +'),
            '+1 +Buk +flat +' + locLabel.replace(' ', ' +'),
            '+1 +Buk +flats +' + locLabel.replace(' ', ' +'),
            '+1 +Buk +apartment +' + locLabel.replace(' ', ' +'),
            '+1 +Buk +apartments +' + locLabel.replace(' ', ' +'),
            '+1 +BHK for +sale in +' + locLabel.replace(' ', ' +'),
            '+' + locLabel.replace(' ', ' +') + ' +1 +BHK ',
            '+' + locLabel.replace(' ', ' +') + ' +1 +BHK +flats',
            '+' + locLabel.replace(' ', ' +') + ' +1 +BHK +apartment',
            '+buy +1 +BHK in +' + locLabel.replace(' ', ' +'),
            '+' + locLabel.replace(' ', ' +') + ' +1 +BHK +apartments'
        ],
        adGroup[7]: [
            '+2 +BHK +' + locLabel.replace(' ', ' +'),
            '+2 +Buk +flat +' + locLabel.replace(' ', ' +'),
            '+2 +Buk +flats +' + locLabel.replace(' ', ' +'),
            '+2 +Buk +apartment +' + locLabel.replace(' ', ' +'),
            '+2 +Buk +apartments +' + locLabel.replace(' ', ' +'),
            '+2 +BHK for +sale in +' + locLabel.replace(' ', ' +'),
            '+' + locLabel.replace(' ', ' +') + ' +2 +BHK ',
            '+' + locLabel.replace(' ', ' +') + ' +2 +BHK +flats',
            '+' + locLabel.replace(' ', ' +') + ' +2 +BHK +apartment',
            '+buy +2 +BHK in +' + locLabel.replace(' ', ' +'),
            '+' + locLabel.replace(' ', ' +') + ' +2 +BHK +apartments'
        ],
        adGroup[8]: [
            '+3 +BHK +' + locLabel.replace(' ', ' +'),
            '+3 +Buk +flat +' + locLabel.replace(' ', ' +'),
            '+3 +Buk +flats +' + locLabel.replace(' ', ' +'),
            '+3 +Buk +apartment +' + locLabel.replace(' ', ' +'),
            '+3 +Buk +apartments +' + locLabel.replace(' ', ' +'),
            '+3 +BHK for +sale in +' + locLabel.replace(' ', ' +'),
            '+' + locLabel.replace(' ', ' +') + ' +3 +BHK ',
            '+' + locLabel.replace(' ', ' +') + ' +3 +BHK +flats',
            '+' + locLabel.replace(' ', ' +') + ' +3 +BHK +apartment',
            '+buy +3 +BHK in +' + locLabel.replace(' ', ' +'),
            '+' + locLabel.replace(' ', ' +') + ' +3 +BHK +apartments'
        ],
        adGroup[9]: [
            '+4 +BHK +' + locLabel.replace(' ', ' +'),
            '+4 +Buk +flat +' + locLabel.replace(' ', ' +'),
            '+4 +Buk +flats +' + locLabel.replace(' ', ' +'),
            '+4 +Buk +apartment +' + locLabel.replace(' ', ' +'),
            '+4 +Buk +apartments +' + locLabel.replace(' ', ' +'),
            '+4 +BHK for +sale in +' + locLabel.replace(' ', ' +'),
            '+' + locLabel.replace(' ', ' +') + ' +4 +BHK ',
            '+' + locLabel.replace(' ', ' +') + ' +4 +BHK +flats',
            '+' + locLabel.replace(' ', ' +') + ' +4 +BHK +apartment',
            '+buy +4 +BHK in +' + locLabel.replace(' ', ' +'),
            '+' + locLabel.replace(' ', ' +') + ' +4 +BHK +apartments'
        ],
        adGroup[10]: [
            '+budget +property in +' + locLabel.replace(' ', ' +'),
            '+budget +properties in +' + locLabel.replace(' ', ' +'),
            '+' + locLabel.replace(' ', ' +') + ' +budget +property',
            '+' + locLabel.replace(' ', ' +') + ' +budget +properties',
            '+budget +property for +sale in +' + locLabel.replace(' ', ' +'),
            '+budget +properties for +sale in +' + locLabel.replace(' ', ' +'),
            '+buy +budget +property  in +' + locLabel.replace(' ', ' +'),
            '+' + locLabel.replace(' ', ' +') + ' +budget +property for +sale',
            '+' + locLabel.replace(' ', ' +') + ' +budget +property'
        ],
        adGroup[11]: [
            '+affordable +property in +' + locLabel.replace(' ', ' +'),
            '+affordable +properties in +' + locLabel.replace(' ', ' +'),
            '+' + locLabel.replace(' ', ' +') + ' +affordable +property',
            '+' + locLabel.replace(' ', ' +') + ' +affordable +properties',
            '+affordable +property for +sale in +' +
            locLabel.replace(' ', ' +'),
            '+affordable +properties for +sale in +' +
            locLabel.replace(' ', ' +'),
            '+buy +affordable +property  in +' + locLabel.replace(' ', ' +'),
            '+' + locLabel.replace(' ', ' +') +
            ' +affordable +property for +sale',
            '+' + locLabel.replace(' ', ' +') + ' +affordable +property'
        ]
    }

    lcUrls = getLastColumnUrl(locId)

    for i, adg in enumerate(adGroup):
        if toBeAdded[adGroup[i]]:
            populateAdsWorksheet(i, adGroup, locLabel, campaign, locId, lcUrls);
            s2_rowNum = s2_rowNum + 1
            for keyString in keyWordsColumn[adGroup[i]]:
                worksheet1.write(s1_rowNum, 0, campaign)
                worksheet1.write(s1_rowNum, 1, adGroup[i])
                worksheet1.write(s1_rowNum, 2, keyString)
                s1_rowNum = s1_rowNum + 1

    s1_rowNum = s1_rowNum + 1
    s2_rowNum = s2_rowNum + 1

scriptInit()
