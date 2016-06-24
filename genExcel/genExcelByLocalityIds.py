#!/usr/bin/python
# Name: genExcelByLocalityIds.py
# Author: [Govindrao Kulkarni]
# Date: 24 June, 2016
# Description: Generate 'Keywords' and 'Ads' Excel WorkSheets for
# Marketing Team

import xlrd
import xlsxwriter
import requests
import json
import ConfigParser
import progressbar as pb

config = ConfigParser.ConfigParser()
config.read("config.cfg")

keep_all_rows = config.get('controlVars', 'keepAllRows')
inputFile = config.get('paths', 'inputFile')
startRow = int(config.get('controlVars', 'startRow'))
endRow = int(config.get('controlVars', 'endRow'))
colNum = int(config.get('controlVars', 'colNum'))

print 'Kepp All Rows: ', keep_all_rows

locIdsList = []
locIdDataMap = {}

exceptionIdList = []

s1_rowNum = 0
s2_rowNum = 0

# Create a output workbook and add a worksheet.
workbook = xlsxwriter.Workbook('marketing.xlsx')
worksheet1 = workbook.add_worksheet('keywords')
worksheet2 = workbook.add_worksheet('ads')

_widgets=[pb.Bar('=', '[', ']'), ' ', pb.Percentage()]
progress = pb.ProgressBar(widgets=_widgets, maxval = 500000).start()
progvar = 0


def exitScript():
    global workbook
    print 'Excel Sheet could not be generated for following locality Ids: ', exceptionIdList
    workbook.close()


def parseLocalityIds():
    global locIdsList
    w1 = xlrd.open_workbook(inputFile)
    w1_s1 = w1.sheet_by_index(0)
    # can do some error checking
    for i in range(startRow, endRow):
        locId = int(w1_s1.cell_value(i, colNum))
        locIdsList.append(locId)


def insert(locId, key, count):
    if locIdDataMap.has_key(locId):
        locIdDataMap[locId][key] = count
    else:
        locIdDataMap[locId] = {}
        locIdDataMap[locId][key] = count


def apiCaller(api):
    print 'GET call: ', api
    res = requests.get(api)
    if res.status_code == requests.codes.ok:
        apiData = json.loads(res.content)
        return apiData
    else:
        print 'ERROR in GET call: ', api
        exitScript()
        res.raise_for_status()
        return None


def populateLabelsInMap(url):
    apiData = apiCaller(url)
    if apiData is not None and apiData.has_key('data'):
        for labelObj in apiData['data']:
            if labelObj.has_key('localityId'):
                locId = str(labelObj['localityId'])
            if labelObj.has_key('suburb') and labelObj['suburb'].has_key('city') and labelObj['suburb']['city'].has_key('label'):
                cityLabel = labelObj['suburb']['city']['label']
                insert(locId, 'cityLabel', cityLabel)
            if labelObj.has_key('label'):
                localityLabel = labelObj['label']
                insert(locId, 'localityLabel', localityLabel)


def populateBhkUrlsInMap(url):
    apiData = apiCaller(url)
    if apiData is not None and apiData.has_key('data'):
        for key in apiData['data']:
            bhkUrl = apiData['data'][key]
            locId = key.replace('MAKAAN_LOCALITY_BHK_PROPERTY_BUY-', '')
            insert(locId, 'bhkUrl', bhkUrl)


def populateListingUrlsInMap(url):
    apiData = apiCaller(url)
    if apiData is not None and apiData.has_key('data'):
        for key in apiData['data']:
            listingUrl = apiData['data'][key]
            locId = key.replace('MAKAAN_LOCALITY_LISTING_BUY-', '')
            insert(locId, 'listingUrl', listingUrl)


def gatherLocalityData():
    apiList1 = {
        'labels': 'https://www.proptiger.com/data/v1/entity/locality?selector={"filters":{"and":[{"equal":{"localityId":' + str(locIdsList) + '}}]},"paging":{"start":0,"rows":10000},"fields":["localityId","suburb","city","label"]}&sourceDomain=Makaan',
        'bhkUrls': 'https://www.makaan.com/dawnstar/data/v2/fetch-urls?urlParam=[{"urlDomain":"locality","domainIds":' + str(locIdsList) + ',"urlCategoryName":"MAKAAN_LOCALITY_BHK_PROPERTY_BUY"}]',
        'listingUrls': 'https://www.makaan.com/dawnstar/data/v2/fetch-urls?urlParam=[{"urlDomain":"locality","domainIds":' + str(locIdsList) + ',"urlCategoryName":"MAKAAN_LOCALITY_LISTING_BUY"}]'
    }
    populateLabelsInMap(apiList1['labels'])
    populateBhkUrlsInMap(apiList1['bhkUrls'])
    populateListingUrlsInMap(apiList1['listingUrls'])

    if keep_all_rows:
        apiList2 = {
            'totalCount': 'http://proptiger.com/app/v1/listing?selector={"filters":{"and":[{"equal":{"listingCategory":["Primary","Resale"]}}]},"paging":{"start":0,"rows":0}}&facets=localityId&sourceDomain=Makaan',
            'bhk1Count': 'http://www.proptiger.com/app/v1/listing?selector={"filters":{"and":[{"equal":{"listingCategory":["Primary","Resale"]}},{"equal":{"bedrooms":["1"]}},{"range":{"price":{"from":"0","to":"5000000"}}}]},"paging":{"start":0,"rows":0}}&facets=localityId&sourceDomain=Makaan',
            'bhk2Count': 'http://www.proptiger.com/app/v1/listing?selector={"filters":{"and":[{"equal":{"listingCategory":["Primary","Resale"]}},{"equal":{"bedrooms":["2"]}},{"range":{"price":{"from":"0","to":"5000000"}}}]},"paging":{"start":0,"rows":0}}&facets=localityId&sourceDomain=Makaan',
            'bhk3Count': 'http://www.proptiger.com/app/v1/listing?selector={"filters":{"and":[{"equal":{"listingCategory":["Primary","Resale"]}},{"equal":{"bedrooms":["3"]}},{"range":{"price":{"from":"0","to":"5000000"}}}]},"paging":{"start":0,"rows":0}}&facets=localityId&sourceDomain=Makaan',
            'bhk4Count': 'http://www.proptiger.com/app/v1/listing?selector={"filters":{"and":[{"equal":{"listingCategory":["Primary","Resale"]}},{"equal":{"bedrooms":["4"]}},{"range":{"price":{"from":"0","to":"5000000"}}}]},"paging":{"start":0,"rows":0}}&facets=localityId&sourceDomain=Makaan'
        }
        for key in apiList2:
            apiData = apiCaller(apiList2[key])
            if apiData.has_key('data') and apiData['data'].has_key('facets') and apiData['data']['facets'].has_key('localityId'):
                locArr = apiData['data']['facets']['localityId']
                for countObj in locArr:
                    for locId in countObj:
                        insert(locId, key, countObj[locId])


def init():
    parseLocalityIds()
    gatherLocalityData()
    for locId in locIdsList:
        generateCurrLocalityContent(locId)
    exitScript()


def getCountFromMap(locId, key):
    return locIdDataMap.has_key(locId) and locIdDataMap[locId].has_key(key) and locIdDataMap[locId][key] > 5


def getToBeAddedData(locId, locLabel):
    return {
        locLabel + ' ' + 'Property': keep_all_rows or getCountFromMap(),
        locLabel + ' ' + 'Apartments': keep_all_rows or getCountFromMap(),
        locLabel + ' ' + 'Homes': keep_all_rows or getCountFromMap(),
        locLabel + ' ' + 'House': keep_all_rows or getCountFromMap(),
        locLabel + ' ' + 'flats': keep_all_rows or getCountFromMap(),
        locLabel + ' ' + 'Builder Floors': keep_all_rows or getCountFromMap(),
        '1BHK' + ' ' + locLabel: keep_all_rows or getCountFromMap(),
        '2BHK' + ' ' + locLabel: keep_all_rows or getCountFromMap(),
        '3BHK' + ' ' + locLabel: keep_all_rows or getCountFromMap(),
        '4BHK' + ' ' + locLabel: keep_all_rows or getCountFromMap(),
        'Budget Proeprty' + ' ' + locLabel: keep_all_rows or getCountFromMap(),
        'Affordable Property' + ' ' + locLabel: keep_all_rows or getCountFromMap()
    }


def populateAdsWorksheet(adgNum, adGroup, locLabel, campaign, locId):
    worksheet2.write(s2_rowNum, 0, campaign)
    if adgNum == 0:
        worksheet2.write(s2_rowNum, 1, adGroup[adgNum])
        worksheet2.write(s2_rowNum, 2, locLabel + ' Property')
        worksheet2.write(s2_rowNum, 3, 'Huge Options in Multiple Budgets.')
        worksheet2.write(s2_rowNum, 4, 'Best Price only on Makaan.com')
        worksheet2.write(s2_rowNum, 5, 'Makaan.com/' +
                         locLabel.replace(" ", "") + '_Property')
        worksheet2.write(s2_rowNum, 6, locIdDataMap[locId]['listingUrl'])
    elif adgNum == 1:
        worksheet2.write(s2_rowNum, 1, adGroup[adgNum])
        worksheet2.write(s2_rowNum, 2, locLabel + ' Apartment')
        worksheet2.write(s2_rowNum, 3, 'Huge Options in Multiple Budgets.')
        worksheet2.write(s2_rowNum, 4, 'Best Price only on Makaan.com')
        worksheet2.write(s2_rowNum, 5, 'Makaan.com/' +
                         locLabel.replace(" ", "") + '_Apts')
        worksheet2.write(s2_rowNum, 6, locIdDataMap[locId]['listingUrl'])
    elif adgNum == 2:
        worksheet2.write(s2_rowNum, 1, adGroup[adgNum])
        worksheet2.write(s2_rowNum, 2, locLabel + ' Home')
        worksheet2.write(s2_rowNum, 3, 'Huge Options in Multiple Budgets.')
        worksheet2.write(s2_rowNum, 4, 'Best Price only on Makaan.com')
        worksheet2.write(s2_rowNum, 5, 'Makaan.com/' +
                         locLabel.replace(" ", "_") + '_Home')
        worksheet2.write(s2_rowNum, 6, locIdDataMap[locId]['listingUrl'])
    elif adgNum == 3:
        worksheet2.write(s2_rowNum, 1, adGroup[adgNum])
        worksheet2.write(s2_rowNum, 2, locLabel + ' House')
        worksheet2.write(s2_rowNum, 3, '100% Real & Verified Properties.')
        worksheet2.write(s2_rowNum, 4, 'Find by Location & Budget Now!')
        worksheet2.write(s2_rowNum, 5, 'Makaan.com/' +
                         locLabel.replace(" ", "_") + '_House')
        worksheet2.write(s2_rowNum, 6, locIdDataMap[locId]['listingUrl'])
    elif adgNum == 4:
        worksheet2.write(s2_rowNum, 1, adGroup[adgNum])
        worksheet2.write(s2_rowNum, 2, 'Flats in ' + locLabel)
        worksheet2.write(s2_rowNum, 3, '100% Real & Verified Properties.')
        worksheet2.write(s2_rowNum, 4, 'Find by Location & Budget Now!')
        worksheet2.write(s2_rowNum, 5, 'Makaan.com/Flats_' +
                         locLabel.replace(" ", "_"))
        worksheet2.write(s2_rowNum, 6, locIdDataMap[locId]['listingUrl'])
    elif adgNum == 5:
        worksheet2.write(s2_rowNum, 1, adGroup[adgNum])
        worksheet2.write(s2_rowNum, 2, locLabel + ' Property')
        worksheet2.write(s2_rowNum, 3, 'Buy Builder Floor, Best Options.')
        worksheet2.write(s2_rowNum, 4, 'Best Price only on Makaan.com')
        worksheet2.write(s2_rowNum, 5, 'Makaan.com/' +
                         locLabel.replace(" ", "") + '_Property')
        worksheet2.write(s2_rowNum, 6, locIdDataMap[locId]['listingUrl'])
    elif adgNum == 6:
        worksheet2.write(s2_rowNum, 1, adGroup[adgNum])
        worksheet2.write(s2_rowNum, 2, locLabel + ' 1 BHK')
        worksheet2.write(s2_rowNum, 3, '100% Real & Verified Properties.')
        worksheet2.write(s2_rowNum, 4, 'Find by Location & Budget Now!')
        worksheet2.write(s2_rowNum, 5, 'Makaan.com/' +
                         locLabel.replace(" ", "_") + '_1_BHK')
        worksheet2.write(s2_rowNum, 6, locIdDataMap[locId][
                         'bhkUrl'].replace('-bhk-', '-1bhk-'))
    elif adgNum == 7:
        worksheet2.write(s2_rowNum, 1, adGroup[adgNum])
        worksheet2.write(s2_rowNum, 2, locLabel + ' 2 BHK')
        worksheet2.write(s2_rowNum, 3, 'Huge Options in Multiple Budgets.')
        worksheet2.write(s2_rowNum, 4, 'Best Price only on Makaan.com')
        worksheet2.write(s2_rowNum, 5, 'Makaan.com/' +
                         locLabel.replace(" ", "_") + '_2_BHK')
        worksheet2.write(s2_rowNum, 6, locIdDataMap[locId][
                         'bhkUrl'].replace('-bhk-', '-2bhk-'))
    elif adgNum == 8:
        worksheet2.write(s2_rowNum, 1, adGroup[adgNum])
        worksheet2.write(s2_rowNum, 2, locLabel + ' 2 BHK')
        worksheet2.write(s2_rowNum, 3, '100% Real & Verified Properties.')
        worksheet2.write(s2_rowNum, 4, 'Find by Location & Budget Now!')
        worksheet2.write(s2_rowNum, 5, 'Makaan.com/' +
                         locLabel.replace(" ", "_") + '_3_BHK')
        worksheet2.write(s2_rowNum, 6, locIdDataMap[locId][
                         'bhkUrl'].replace('-bhk-', '-3bhk-'))
    elif adgNum == 9:
        worksheet2.write(s2_rowNum, 1, adGroup[adgNum])
        worksheet2.write(s2_rowNum, 2, locLabel + ' 4 BHK')
        worksheet2.write(s2_rowNum, 3, 'Huge Options in Multiple Budgets.')
        worksheet2.write(s2_rowNum, 4, 'Best Price only on Makaan.com')
        worksheet2.write(s2_rowNum, 5, 'Makaan.com/' +
                         locLabel.replace(" ", "_") + '_4_BHK')
        worksheet2.write(s2_rowNum, 6, locIdDataMap[locId][
                         'bhkUrl'].replace('-bhk-', '-4bhk-'))
    elif adgNum == 10:
        worksheet2.write(s2_rowNum, 1, adGroup[adgNum])
        worksheet2.write(s2_rowNum, 2, locLabel + ' Property')
        worksheet2.write(s2_rowNum, 3, 'Property for Every Budget.')
        worksheet2.write(s2_rowNum, 4, 'Best Price only on Makaan.com')
        worksheet2.write(s2_rowNum, 5, 'Makaan.com/' +
                         locLabel.replace(" ", "") + '_Property')
        worksheet2.write(s2_rowNum, 6, locIdDataMap[locId]['listingUrl'])
    elif adgNum == 11:
        worksheet2.write(s2_rowNum, 1, adGroup[adgNum])
        worksheet2.write(s2_rowNum, 2, locLabel + ' Property')
        worksheet2.write(s2_rowNum, 3, 'Wide Range of Affordable Properties')
        worksheet2.write(s2_rowNum, 4, 'Find by Location & Budget Now!')
        worksheet2.write(s2_rowNum, 5, 'Makaan.com/' +
                         locLabel.replace(" ", "") + '_Property')
        worksheet2.write(s2_rowNum, 6, locIdDataMap[locId]['listingUrl'])


def generateCurrLocalityContent(locId):
    global worksheet1
    global s1_rowNum
    global s2_rowNum

    locId = str(locId)
    print locIdDataMap[locId]
    if locIdDataMap.has_key(locId) and locIdDataMap[locId].has_key('cityLabel'):
        cityLabel = locIdDataMap[locId]['cityLabel']
    else:
        print 'Error, could not get cityLabel: ', locId
        exceptionIdList.append(locId)
        return
    if locIdDataMap.has_key(locId) and locIdDataMap[locId].has_key('localityLabel'):
        locLabel = locIdDataMap[locId]['localityLabel']
    else:
        print 'Error, could not get localityLabel: ', locId
        exceptionIdList.append(locId)
        return

    campaign = cityLabel + '-' + locLabel
    toBeAdded = getToBeAddedData(locId, locLabel)

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

    for i, adg in enumerate(adGroup):
        if toBeAdded[adGroup[i]]:
            populateAdsWorksheet(i, adGroup, locLabel, campaign, locId)
            s2_rowNum = s2_rowNum + 1
            for keyString in keyWordsColumn[adGroup[i]]:
                worksheet1.write(s1_rowNum, 0, campaign)
                worksheet1.write(s1_rowNum, 1, adGroup[i])
                worksheet1.write(s1_rowNum, 2, keyString)
                s1_rowNum = s1_rowNum + 1

    s1_rowNum = s1_rowNum + 1
    s2_rowNum = s2_rowNum + 1


for i in range(500000):
    init()
    progress.update(progvar + 1)
    progvar += 1