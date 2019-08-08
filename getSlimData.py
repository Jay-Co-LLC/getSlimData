import os
import datetime
import requests
import json
import logging
import ftplib
from threading import Thread
from threading import RLock
import openpyxl as XL
import xml.etree.ElementTree as ET

logger = logging.getLogger()
logger.setLevel(logging.ERROR)

url = 'https://api.ebay.com/ws/api.dll'

userid = os.environ['userid']
key = os.environ['key']
ftp_user = os.environ['ftp_user']
ftp_pass = os.environ['ftp_pass']

getAllItemIdsParams = {
	'Content-Type' : 'text/xml',
	'X-EBAY-API-COMPATIBILITY-LEVEL' : '1081',
	'X-EBAY-API-CALL-NAME' : 'GetSellerList',
	'X-EBAY-API-SITEID' : '0'
	}
	
getAllItemsParams = {
	'Content-Type' : 'text/xml',
	'X-EBAY-API-COMPATIBILITY-LEVEL' : '1081',
	'X-EBAY-API-CALL-NAME' : 'GetItem',
	'X-EBAY-API-SITEID' : '0'
	}

today = datetime.datetime.now() - datetime.timedelta(hours=7)
future = today + datetime.timedelta(days=120)

shortnames = {
	'suspensionspecialists' : 'TSS',
	'wulfsuspension' : 'WLF'
	}

def P(str):
	return f"{{urn:ebay:apis:eBLBaseComponents}}{str}"
	
def getValueString(name, item):
	itemspecifics = item.find(P('ItemSpecifics'))
	returnString = ''
	
	for each in itemspecifics:
		if (each.find(P('Name')).text == name):
			allValues = each.findall(P('Value'))
			numValues = len(allValues)
			
			if (numValues > 1):
				i = 0
				while (i < numValues):
					if (i != (numValues - 1)):
						returnString = allValues[i].text + '|' + returnString
					else:
						returnString = returnString + allValues[i].text
					i = i + 1
			else:
				returnString = allValues[0].text
				break
	
	return returnString

def getAllItemIdsXML(pagenum):
	return f"""
<?xml version="1.0" encoding="utf-8"?>
<GetSellerListRequest xmlns="urn:ebay:apis:eBLBaseComponents">
  <RequesterCredentials>  
  <eBayAuthToken>{key}</eBayAuthToken>
  </RequesterCredentials>
  <EndTimeFrom>{today}</EndTimeFrom>
  <EndTimeTo>{future}</EndTimeTo>
  <Pagination>
    <EntriesPerPage>200</EntriesPerPage>
    <PageNumber>{pagenum}</PageNumber>
  </Pagination>
  <OutputSelector>ItemID</OutputSelector>
  <OutputSelector>PaginationResult</OutputSelector>
</GetSellerListRequest>
"""

def getAllItemsXML(itemid):
	return f"""
<?xml version="1.0" encoding="utf-8"?>
<GetItemRequest xmlns="urn:ebay:apis:eBLBaseComponents">
  <RequesterCredentials>
    <eBayAuthToken>{key}</eBayAuthToken>
  </RequesterCredentials>
  <ItemID>{itemid}</ItemID>
  <IncludeItemSpecifics>True</IncludeItemSpecifics>
  <DetailLevel>ReturnAll</DetailLevel>
</GetItemRequest>
"""

def getAllItemIds():
	cur_page = 1
	tot_pages = 1

	itemids = []
	
	logger.info(f"[{userid}] Starting getAllItemIds...")
	
	while (cur_page <= tot_pages):
		logger.info(f"[{userid}] Calling GetSellerList - Page {cur_page}")
		r = requests.post(url, data=getAllItemIdsXML(cur_page), headers=getAllItemIdsParams)
		
		logger.info(f"[{userid}] Response code: {r.status_code}")
		
		if (r.status_code != 200):
			logger.error(f"[{userid}] Response: {r.text}")
		
		root = ET.fromstring(r.content)
	
		tot_pages = int(root.find(P('PaginationResult')).find(P('TotalNumberOfPages')).text)
	
		itemArr = root.find(P('ItemArray'))
	
		for eachItem in itemArr:
			itemid = eachItem.find(P('ItemID')).text
			itemids.append(itemid)
		
		# Append each half of the list of 200, we want to create 1 thread per 100 listings
		allItemIds.append(itemids[:len(itemids)//2])
		allItemIds.append(itemids[len(itemids)//2:])
		
		itemids = []
		cur_page = cur_page + 1

def getItems(listOfItemIds):
	for eachItemId in listOfItemIds:
		logger.info(f"[{userid}] Calling GetItem for {eachItemId}")
		r = requests.post(url, data=getAllItemsXML(eachItemId), headers=getAllItemsParams)
		
		logger.info(f"[{userid}] Response code: {r.status_code}")
		
		if (r.status_code != 200):
			logger.error(f"[{userid}] Response: {r.text}")
		
		root = ET.fromstring(r.content)
		item = root.find(P('Item'))

		CategoryID = ''
		try:
			CategoryID = item.find(P('PrimaryCategory')).find(P('CategoryID')).text
		except:
			pass
		
		StoreCategoryID = ''
		try:
			StoreCategoryID = item.find(P('Storefront')).find(P('StoreCategoryID')).text
		except:
			pass
		
		Title = ''
		try:
			Title = item.find(P('Title')).text
		except:
			pass

		ConditionID = ''
		try:
			ConditionID = item.find(P('ConditionID')).text
		except:
			pass

		Brand = getValueString('Brand', item)	
		PartType = getValueString('Part Type', item)
		ManufacturerPartNumber = getValueString('Manufacturer Part Number', item)
		InterchangePartNumber = getValueString('Interchange Part Number', item)
		OtherPartNumber = getValueString('Other Part Number', item)
		PlacementOnVehicle = getValueString('Placement on Vehicle', item)
		Warranty = getValueString('Warranty', item)
		CustomBundle = getValueString('Custom Bundle', item)
		FitmentType = getValueString('Fitment Type', item)
		IncludedHardware = getValueString('Included Hardware', item)
		Greasable = getValueString('Greasable', item)
		ModifiedItem = getValueString('Modified Item', item)
		Adjustable = getValueString('Adjustable', item)
		NonDomesticProduct = getValueString('Non-Domestic Product', item)
		CountryRegionOfManufacture = getValueString('Country/Region of Manufacture', item)
		
		PicURL = ''
		try:
			PicURL = item.find(P('PictureDetails')).find(P('GalleryURL')).text
		except:
			pass

		GalleryType = ''
		try:
			GalleryType = item.find(P('PictureDetails')).find(P('GalleryType')).text
		except:
			pass
			
		Description = ''
		try:
			Description = item.find(P('Description')).text
		except:
			pass

		Format = ''

		Duration = ''
		try:
			Duration = item.find(P('ListingDuration')).text
		except:
			pass
			
		StartPrice = ''
		try:
			StartPrice = item.find(P('StartPrice')).text
		except:
			pass
			
		BuyItNowPrice = ''
		try:
			BuyItNowPrice = item.find(P('BuyItNowPrice')).text
		except:
			pass
			
		Quantity = ''
		try:
			Quantity = item.find(P('Quantity')).text
		except:
			pass
		
		ShippingType = ''
		try:
			ShippingType = item.find(P('ShippingDetails')).find(P('ShippingType')).text
		except:
			pass
		
		ShippingService1Option = ''
		try:
			ShippingService1Option = item.find(P('ShippingDetails')).find(P('ShippingServiceOptions')).find(P('ShippingService')).text
		except:
			pass
			
		ShippingService1Cost = ''	
		try:
			ShippingService1Cost = item.find(P('ShippingDetails')).find(P('ShippingServiceOptions')).find(P('ShippingServiceCost')).text
		except:
			pass

		DispatchTimeMax = ''
		try:
			DispatchTimeMax = item.find(P('DispatchTimeMax')).text
		except:
			pass

		ReturnsAcceptedOption = ''
		try:
			ReturnsAcceptedOption = item.find(P('ReturnPolicy')).find(P('ReturnsAcceptedOption')).text
		except:
			pass
			
		ReturnsWithinOption = ''
		try:
			ReturnsWithinOption = item.find(P('ReturnPolicy')).find(P('ReturnsWithinOption')).text
		except:
			pass

		toAdd = {
			'itemid' : eachItemId,
			'Listing URL' : f'https://www.ebay.com/itm/{eachItemId}',
			'*Category' : CategoryID,
			'StoreCategory' : StoreCategoryID,
			'*Title' : Title,
			'*ConditionID' : ConditionID,
			'*C:Brand' : Brand,
			'C:Part Type' : PartType,
			'*C:Manufacturer Part Number' : ManufacturerPartNumber,
			'C:Interchange Part Number' : InterchangePartNumber,
			'C:Other Part Number' : OtherPartNumber,
			'C:Placement on Vehicle' : PlacementOnVehicle,
			'C:Warranty' : Warranty,
			'C:Custom Bundle' : CustomBundle,
			'C:Fitment Type' : FitmentType,
			'C:Included Hardware' : IncludedHardware,
			'C:Greasable or Sealed' : Greasable,
			'C:Modified Item' : ModifiedItem,
			'C:Adjustable' : Adjustable,
			'C:Non-Domestic Product' : NonDomesticProduct,
			'C:Country/Region of Manufacture' : CountryRegionOfManufacture,
			'PicURL' : PicURL,
			'GalleryType' : GalleryType,
			'*Description' : Description,
			'*Format' : Format,
			'*Duration' : Duration,
			'*StartPrice' : StartPrice,
			'BuyItNowPrice' : BuyItNowPrice,
			'*Quantity' : Quantity,
			'ShippingType' : ShippingType,
			'ShippingService-1:Option' : ShippingService1Option,
			'ShippingService-1:Cost' : ShippingService1Cost,
			'*DispatchTimeMax' : DispatchTimeMax,
			'*ReturnsAcceptedOption' : ReturnsAcceptedOption,
			'ReturnsWithinOption' : ReturnsWithinOption}
			
		wbLock.acquire()
		
		try:
			logger.info(f"[{userid}] Writing {toAdd['itemid']} to file")
			outws.append([value for key, value in toAdd.items()])
		finally:
			wbLock.release()
			
outwb = XL.Workbook()
outws = outwb.active

wbLock = RLock()

# Write headers to excel file
outws.append([
	'itemid',
	'Listing URL',
	'*Category',
	'StoreCategory',
	'*Title',
	'*ConditionID',
	'*C:Brand',
	'C:Part Type',
	'*C:Manufacturer Part Number',
	'C:Interchange Part Number',
	'C:Other Part Number',
	'C:Placement on Vehicle',
	'C:Warranty',
	'C:Custom Bundle',
	'C:Fitment Type',
	'C:Included Hardware',
	'C:Greasable or Sealed',
	'C:Modified Item',
	'C:Adjustable',
	'C:Non-Domestic Product',
	'C:Country/Region of Manufacture',
	'PicURL',
	'GalleryType',
	'*Description',
	'*Format',
	'*Duration',
	'*StartPrice',
	'BuyItNowPrice',
	'*Quantity',
	'ShippingType',
	'ShippingService-1:Option',
	'ShippingService-1:Cost',
	'*DispatchTimeMax',
	'*ReturnsAcceptedOption',
	'ReturnsWithinOption'])

threads = []
allItemIds = []

def main(event, context):
	
	getAllItemIds()
		
	for listOfItemIds in allItemIds:
		t = Thread(target=getItems, args=(listOfItemIds,))
		threads.append(t)
		
	for t in threads:
		t.start()
		
	for t in threads:
		t.join()
		
	# Write file to temp local storage
	logger.info(f"[{userid}] Writing file to /tmp/")
	outwb.save('/tmp/out.xlsx')
	
	# Put file to FTP server
	logger.info(f"[{userid}] Copying file from /tmp/ to FTP")
	filename = f"{userid}-data.xlsx"
	
	logger.info(f"[{userid}] Connecting to FTP...")
	
	ftp = ftplib.FTP()
	ftp.connect(os.environ['ftp_ip'], 21, timeout=120)
	ftp.set_debuglevel(1)
	ftp.set_pasv(True)
	ftp.login('sam@jaycotss.com', 'TSSconnect1!')
	
	f = open('/tmp/out.xlsx', 'rb')
	
	logger.info(f"[{userid}] Sending local file to FTP...")
	ftp.storbinary(f'STOR {filename}', f)
	
	f.close()
	
	logger.info(f"[{userid}] Closing FTP session...")
	ftp.quit()
	
	logger.info(f"[{userid}] Done.")

