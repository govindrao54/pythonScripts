#!/usr/bin/python
import requests
import json
print ("Hello, Python!")

exceptionIdList = []

def getLastColumnUrl(locId, adgNum):
    if adgNum in [6, 7, 8, 9]:
        api1 = 'https://www.makaan.com/dawnstar/data/v2/fetch-urls?urlParam=[{"urlDomain":"locality","domainIds":[' + `locId` + '],"urlCategoryName":"MAKAAN_LOCALITY_BHK_PROPERTY_BUY"}]'
        res = requests.get(api1)
        if res.status_code == requests.codes.ok:
            apiData = json.loads(res.content)
            if apiData.has_key('data') and apiData['data'] is not None and apiData['data'].has_key('MAKAAN_LOCALITY_BHK_PROPERTY_BUY-' + `locId`):
                return (apiData['data']['MAKAAN_LOCALITY_BHK_PROPERTY_BUY-' + `locId`]).replace('-bhk-', '-1bhk-')
            else:
                return exceptionIdList.append(locId)
                print 'error in fetching propertyUrl'
        else:
            print 'error in GET call for', api1
            return exceptionIdList.append(locId)
            # If response code is not ok (200), print the resulting http error code with description
            res.raise_for_status()
    elif adgNum in [0, 1, 2, 3, 4, 5, 10, 11]:
        api2 = 'https://www.makaan.com/dawnstar/data/v2/fetch-urls?urlParam=[{"urlDomain":"locality","domainIds":[' + `locId` + '],"urlCategoryName":"MAKAAN_LOCALITY_LISTING_BUY"}]'
        res = requests.get(api2)
        if res.status_code == requests.codes.ok:
            apiData = json.loads(res.content)
            if apiData.has_key('data') and apiData['data'] is not None and apiData['data'].has_key('MAKAAN_LOCALITY_LISTING_BUY-' + `locId`):
                return apiData['data']['MAKAAN_LOCALITY_LISTING_BUY-' + `locId`]
            else:
                return exceptionIdList.append(locId)
                print 'error in fetching propertyBhkUrl'
        else:
            print 'error in GET call for', api2
            return exceptionIdList.append(locId)
            # If response code is not ok (200), print the resulting http error code with description
            res.raise_for_status()


print getLastColumnUrl(50186, 6)
print '____________________________________________ASDF__________________________________'
print getLastColumnUrl(50186, 3)