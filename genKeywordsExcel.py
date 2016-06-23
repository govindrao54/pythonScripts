#!/usr/bin/python
import xlsxwriter
import requests
import json

locIdMap = {
    'Delhi': [
        51195,
        51801
    ],
    'Mumbai': [
        50063,
        50003,
        50412
    ],
    'Ghaziabad': [
        50230,
        51167
    ],
    'Pune': [
        50092,
        50100,
        50099
    ],
    'Chandigharh': [
        51054
    ],
    'Bangalore': [
        54337
    ]
}

locLabelMap = {
    51195: 'Uttam Nagar',
    50063: 'Thane West',
    50003: 'Kharghar',
    50230: 'Indirapuram',
    50092: 'Kharadi',
    50100: 'Wagholi',
    50412: 'Ulwe',
    51054: 'Zirakpur',
    51801: 'vasundhara',
    54337: 'Whitefield',
    50099: 'Viman Nagar',
    51167: 'Raj Nagar Extension'
}

rowNum = 1


def scriptInit():
	global rowNum
    # Create a workbook and add a worksheet.
	workbook = xlsxwriter.Workbook('keyWords.xlsx')
    # Add a bold format to use to highlight cells.
	bold = workbook.add_format({'bold': True})
	for cityLabel in locIdMap:
		worksheet = workbook.add_worksheet(cityLabel)
		worksheet.write('A1', 'Campaign', bold)
		worksheet.write('B1', 'AdGroups', bold)
		worksheet.write('C1', 'Keywords', bold)
		rowNum = 1
		for locId in locIdMap[cityLabel]:
			appendCurrLocalityContent(cityLabel, locId, worksheet)
	workbook.close()


def getFacetsAPI(locId):
    url = 'https://www.proptiger.com/app/v2/locality/' + `locId` + '?selector={%22fields%22:[%22localityId%22,%22label%22]}'
    return url


def appendCurrLocalityContent(cityLabel, locId, worksheet):
    res = requests.get(getFacetsAPI(locId))
    if res.status_code == requests.codes.ok:
        locApiData = json.loads(res.content)
        print locApiData
        print 'status OK'
        generateCurrLocalityContent(cityLabel, locId, locApiData, worksheet)
    else:
        print 'error in GET call for', locId
        # If response code is not ok (200), print the resulting http error code
        # with description
        myResponse.raise_for_status()


def generateCurrLocalityContent(cityLabel, locId, locApiData, worksheet):
    locLabel = locLabelMap[locId]
    campaign = cityLabel + '-' + locLabel

    global rowNum

    # True if it should be present
    toBeAdded = {
        locLabel + ' ' + 'Property': True,
        locLabel + ' ' + 'Apartments': True,
        locLabel + ' ' + 'Homes': True,
        locLabel + ' ' + 'House': True,
        locLabel + ' ' + 'flats': True,
        locLabel + ' ' + 'Builder Floors': True,
        '1BHK' + ' ' + locLabel: True,
        '2BHK' + ' ' + locLabel: True,
        '3BHK' + ' ' + locLabel: True,
        '4BHK' + ' ' + locLabel: True,
        'Budget Proeprty' + ' ' + locLabel: True,
        'Affordable Property' + ' ' + locLabel: True
    }

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
            for keyString in keyWordsColumn[adGroup[i]]:
                worksheet.write(rowNum, 0, campaign)
                worksheet.write(rowNum, 1, adGroup[i])
                worksheet.write(rowNum, 2, keyString)
                rowNum = rowNum + 1

    rowNum = rowNum + 1
    print rowNum


scriptInit()
