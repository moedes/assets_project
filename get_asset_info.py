import requests
import requests.exceptions
from requests_ntlm import HttpNtlmAuth
import pandas as pd
import credential
import os
#from datetime import datetime
#import dateutil.parser
#from dateutil.parser import parse, relativedelta

from datetime import datetime
from dateutil.parser import parse
from dateutil.relativedelta import relativedelta


def getSiteID_bySN(sn):
   session = requests.Session()
   session.auth = HttpNtlmAuth(credential.login['username'],credential.login['password'], session)

   r = session.get('http://sitelookup.corp.emc.com/models/getdataJH.php?fld=TLA%20SN&val=' + sn)
   
   if int(r.json()["records"]) > 0:
      data = r.json()
      siteID = data['rows'][0]['Party Number']
      
   return siteID

def getSiteInfo_bySN(sn):
   session = requests.Session()
   session.auth = HttpNtlmAuth(credential.login['username'],credential.login['password'], session)

   r = session.get('http://sitelookup.corp.emc.com/models/getdataJH.php?fld=TLA%20SN&val=' + sn)
   
   if int(r.json()["records"]) > 0:
      data = r.json()
      row = 0
      labels = list(data['rows'][0].keys())
      inventory_df = pd.DataFrame(index=range(0,len(data['rows'])), columns=labels)
      for element in data['rows']:
         for key,value in element.items():
            inventory_df.loc[row,key] = value
         row += 1
   
   siteInfo_dict = inventory_df.to_dict()
      
   return siteInfo_dict

def getAssetInfo_bySN(sn):
      
   session = requests.Session()
   session.auth = HttpNtlmAuth(credential.login['username'],credential.login['password'], session)
   r = session.get('http://sitelookup.corp.emc.com/models/getdataTLA.php?fld=TLA&val=' + sn)
   
   if int(r.json()["records"]) > 0:
      data = r.json()
      row = 0
      labels = list(data['rows'][0].keys())
      inventory_df = pd.DataFrame(index=range(0,len(data['rows'])), columns=labels)
      for element in data['rows']:
         for key,value in element.items():
            inventory_df.loc[row,key] = value
         row += 1
   
   asset_dict = inventory_df.loc[inventory_df['ITEM_SERIAL_NUMBER'] == sn].to_dict()
   
   return asset_dict

def getMaintenanceVisioText(sn):
   
   asset_data = getAssetInfo_bySN(sn)
 
   contractStartDate = parse(asset_data['CONTRACT_SUBLINE_START_DATE'][0])
   contractStartDateText = contractStartDate.strftime("%B %d, %Y")
   
   contractEndDate = parse(asset_data['CONTRACT_SUBLINE_END_DATE'][0])
   contractEndDateText = contractEndDate.strftime("%B %d, %Y")
   
   contractRemainingTime = relativedelta(contractEndDate, datetime.now())
   
   yearsRemaining = contractRemainingTime.years
   monthsRemaining = contractRemainingTime.months
   daysRemaining = contractRemainingTime.days
   
   maint_text = 'Maintenance Contract\nStarted %s\nEnds %s\n%s years and %s months and %s days remaining as of %s' %(contractStartDateText,contractEndDateText,yearsRemaining,monthsRemaining,daysRemaining,datetime.today().strftime('%m-%d-%y'))
   
   return maint_text

def getAssetInfo_bySite(SiteID):

   session = requests.Session()
   session.auth = HttpNtlmAuth(credential.login['username'],credential.login['password'], session)

   r = session.get('http://sitelookup.corp.emc.com/models/getdataTLA.php?fld=Party%20ID&val=' + SiteID)
   
   if int(r.json()["records"]) > 0:
      data = r.json()
      row = 0
      labels = list(data['rows'][0].keys())
      inventory_df = pd.DataFrame(index=range(0,len(data['rows'])), columns=labels)
      for element in data['rows']:
         for key,value in element.items():
            inventory_df.loc[row,key] = value
         row += 1
   else:
      labels = ['CS_CUSTOMER_NAME', 'PARTY_NUMBER', 'GLOBAL_DUNS_NUMBER', 'GLOBAL_DUNS_NAME', 'ITEM_SERIAL_NUMBER', 'MODEL_UNIQUE_IDENTIFIER', 'ITEM_INSTALL_DATE', 'MODEL', 'ITEM_DESCRIPTION', 'ITEM_NUM', 'PRODUCT_GROUP', 'PRODUCT_TYPE', 'PRODUCT_FAMILY', 'INSTANCE_PRODUCT_FAMILY', 'INSTALL_BASE_STATUS', 'Instance Description', 'MICROCODE', 'MAINTAINED_BY_GROUP', 'SERVICE_PROVIDER', 'CONNECT_IN_TYPE', 'CONNECT_HOME_TYPE', 'CONNECTED_TO_SN', 'SYR_LAST_DIAL_HOME_DATE', 'SALES_ORDER', 'SALES_ORDER_TYPE', 'CONTRACT_NUMBER', 'COVERAGE_TYPE', 'CONTRACT_SUBLINE_STATUS', 'CONTRACT_SUBLINE_START_DATE', 'CONTRACT_SUBLINE_END_DATE', 'INTERNAL_CUSTOMER', 'PDR', 'SDR', 'IB Solution', 'VCE Support', 'G Code', 'EH SP', 'Address1', 'Address2', 'City', 'State', 'Province', 'Postal Code', 'Time Zone Name', 'DSM_EMAIL', 'DISTRICT', 'PRIMARY_CE_EMAIL', 'ASR_EMAIL', 'CS_ADVOCATE_EMAIL', 'SAM_EMAIL', 'REGION', 'DIVISION', 'THEATER', 'solutionId', 'solutionName']       
      inventory_df = pd.DataFrame(columns=labels)


   print(len(inventory_df))
   siteAssetInfo = inventory_df
   
   return siteAssetInfo

def getAssetInfo_byDuns(DunsID):

   session = requests.Session()
   session.auth = HttpNtlmAuth(credential.login['username'],credential.login['password'], session)

   r = session.get('http://opsconsole.corp.emc.com/sitelookup/models/getdataTLA.php?fld=Global%20Duns%20Number&val=' + DunsID)
   
   if int(r.json()["records"]) > 0:
      data = r.json()
      row = 0
      labels = list(data['rows'][0].keys())
      inventory_df = pd.DataFrame(index=range(0,len(data['rows'])), columns=labels)
      for element in data['rows']:
         for key,value in element.items():
            inventory_df.loc[row,key] = value
         row += 1
   else:
      labels = ['CS_CUSTOMER_NAME', 'PARTY_NUMBER', 'GLOBAL_DUNS_NUMBER', 'GLOBAL_DUNS_NAME', 'ITEM_SERIAL_NUMBER', 'MODEL_UNIQUE_IDENTIFIER', 'ITEM_INSTALL_DATE', 'MODEL', 'ITEM_DESCRIPTION', 'ITEM_NUM', 'PRODUCT_GROUP', 'PRODUCT_TYPE', 'PRODUCT_FAMILY', 'INSTANCE_PRODUCT_FAMILY', 'INSTALL_BASE_STATUS', 'Instance Description', 'MICROCODE', 'MAINTAINED_BY_GROUP', 'SERVICE_PROVIDER', 'CONNECT_IN_TYPE', 'CONNECT_HOME_TYPE', 'CONNECTED_TO_SN', 'SYR_LAST_DIAL_HOME_DATE', 'SALES_ORDER', 'SALES_ORDER_TYPE', 'CONTRACT_NUMBER', 'COVERAGE_TYPE', 'CONTRACT_SUBLINE_STATUS', 'CONTRACT_SUBLINE_START_DATE', 'CONTRACT_SUBLINE_END_DATE', 'INTERNAL_CUSTOMER', 'PDR', 'SDR', 'IB Solution', 'VCE Support', 'G Code', 'EH SP', 'Address1', 'Address2', 'City', 'State', 'Province', 'Postal Code', 'Time Zone Name', 'DSM_EMAIL', 'DISTRICT', 'PRIMARY_CE_EMAIL', 'ASR_EMAIL', 'CS_ADVOCATE_EMAIL', 'SAM_EMAIL', 'REGION', 'DIVISION', 'THEATER', 'solutionId', 'solutionName']       
      inventory_df = pd.DataFrame(columns=labels)
        
   
   print(len(inventory_df))
   dunsAssetInfo = inventory_df
   
   return dunsAssetInfo

def getAssetInfo_forSites(siteID_List):
       
    for index, siteID in enumerate(siteID_List):
        
        print('index %s, site id %s' %(index, siteID))
        
        if index == 0:
            print(siteID)
            assetBook = getAssetInfo_bySite(siteID)
        else:
            print(siteID)
            assetBook = pd.concat([assetBook, getAssetInfo_bySite(siteID)], ignore_index=True)  
    
    return assetBook

def getAssetInfo_forDunsList(dunsID_List):
       
    for index, dunsID in enumerate(dunsID_List):
        
        if index == 0:
            print(dunsID)
            assetBook = getAssetInfo_byDuns(dunsID)
        else:
            print(dunsID)
            assetBook = pd.concat([assetBook, getAssetInfo_byDuns(dunsID)], ignore_index=True)  
    
    labels = ['CS_CUSTOMER_NAME', 'PARTY_NUMBER', 'GLOBAL_DUNS_NUMBER', 'GLOBAL_DUNS_NAME', 'ITEM_SERIAL_NUMBER', 'MODEL_UNIQUE_IDENTIFIER', 'ITEM_INSTALL_DATE', 'MODEL', 'ITEM_DESCRIPTION', 'ITEM_NUM', 'PRODUCT_GROUP', 'PRODUCT_TYPE', 'PRODUCT_FAMILY', 'INSTANCE_PRODUCT_FAMILY', 'INSTALL_BASE_STATUS', 'Instance Description', 'MICROCODE', 'MAINTAINED_BY_GROUP', 'SERVICE_PROVIDER', 'CONNECT_IN_TYPE', 'CONNECT_HOME_TYPE', 'CONNECTED_TO_SN', 'SYR_LAST_DIAL_HOME_DATE', 'SALES_ORDER', 'SALES_ORDER_TYPE', 'CONTRACT_NUMBER', 'COVERAGE_TYPE', 'CONTRACT_SUBLINE_STATUS', 'CONTRACT_SUBLINE_START_DATE', 'CONTRACT_SUBLINE_END_DATE', 'INTERNAL_CUSTOMER', 'PDR', 'SDR', 'IB Solution', 'VCE Support', 'G Code', 'EH SP', 'Address1', 'Address2', 'City', 'State', 'Province', 'Postal Code', 'Time Zone Name', 'DSM_EMAIL', 'DISTRICT', 'PRIMARY_CE_EMAIL', 'ASR_EMAIL', 'CS_ADVOCATE_EMAIL', 'SAM_EMAIL', 'REGION', 'DIVISION', 'THEATER', 'solutionId', 'solutionName']
    assetBook = assetBook[labels]
    
    return assetBook

sn = 'HK197000648'
maintVisioText = getMaintenanceVisioText(sn)
print(maintVisioText)

siteID = getSiteID_bySN(sn)
#siteID = '38545'

assetBook = getAssetInfo_bySite(siteID)
savepath = os.path.abspath(r'C:\Users\cohend\Documents\before\projects MIT\Focus on the Family\Assessment\X1')
writer = pd.ExcelWriter(os.path.join(savepath, 'Assets - Site ID ' + siteID + '.xlsx'))
assetBook.to_excel(writer,'asset info')
writer.save()
writer.close()

# UTAH
dunsList = ['618510226', '363816083', '007939176', '079364201', '032601364', '008954489', '009079831', '625922679', '045807195', '047114848', '009428988', '006813109', '055886908', '187557731', '063560986', '968628151', '869415658', '142606370', '949301337', '137293051', '832274935', '079270980', '196045694', '031861203', '617191122', '808265917', '008523374', '045263035', '051299258', '958668055', '967343869', '028931728', '063297154', '142429468', '786195701', '793506390', '010620755', '073111676', '790845171', '807419051', '945968246', '078659396', '177861908', '145093139', '160122222', '121214985', '126930226', '148453400', '027296810', '129673443', '119732696', '080159438', '829460117', '007597313', '790657378', '809528300', '040760006', '177932639', '847837890', '073115537', '611204835', '035375757', '966817975', '001947688', '056090319', '011266280', '843239984', '007875164', '079824528', '800970592', '623932451', '035315621', '102704954', '077535490', '047112560', '801927492', '800101763', '088188821', '783359602', '010638241', '080041801', '831613646', '063311922', '187059332', '015692325', '809153971', '804413250', '796192243', '080334434', '111339649', '044719144', '125627120']
assetBook = getAssetInfo_forDunsList(dunsList)
savepath = os.path.abspath(r'C:\Users\cohend\Documents\work data MDC\Administration\Commercial')
writer = pd.ExcelWriter(os.path.join(savepath, 'Assets - UTAH ' + '.xlsx'))
assetBook.to_excel(writer,'asset info')
writer.save()
writer.close()

# State of CO OIT EFORT
siteID_List = ['13468438', '13563174']
assetBook = getAssetInfo_forSites(siteID_List)
savepath = os.path.abspath(r'C:\Users\cohend\Documents')
writer = pd.ExcelWriter(os.path.join(savepath, 'Assets - OIT EFORT ' + '-'.join(siteID_List) + '.xlsx'))
assetBook.to_excel(writer,'asset info')
writer.save()
writer.close()

# State of CO OIT KIPLING
siteID_List = ['3672823', '14362892']
assetBook = getAssetInfo_forSites(siteID_List)
savepath = os.path.abspath(r'C:\Users\cohend\Documents\projects MDC\State of CO OIT\Asset Info')
writer = pd.ExcelWriter(os.path.join(savepath, 'Assets - OIT KIPLING ' + '-'.join(siteID_List) + '.xlsx'))
assetBook.to_excel(writer,'asset info')
writer.save()
writer.close()

# UCHealth
siteID_List = ['4038878', '1003871014']
assetBook = getAssetInfo_forSites(siteID_List)
savepath = os.path.abspath(r'C:\Users\cohend\Documents\projects MDC\UCH\Assessments\Asset Info')
writer = pd.ExcelWriter(os.path.join(savepath, 'Assets - UCHealth ' + '-'.join(siteID_List) + '.xlsx'))
assetBook.to_excel(writer,'asset info')
writer.save()
writer.close()

# DEN - Airport
siteID_List = ['1003855654','38545', '3939618', '14584553']
assetBook = getAssetInfo_forSites(siteID_List)
savepath = os.path.abspath(r'C:\Users\cohend\Documents\projects MDC\DEN\Assessments\Asset Info')
writer = pd.ExcelWriter(os.path.join(savepath, 'Assets - DEN ' + '-'.join(siteID_List) + '.xlsx'))
assetBook.to_excel(writer,'asset info')
writer.save()
writer.close()

# CCOD
siteID_List = ['1003925225', '2530711', '2530727', '26349867', '3986941', '663354']
assetBook = getAssetInfo_forSites(siteID_List)
savepath = os.path.abspath(r'C:\Users\cohend\Documents\projects MDC\State of CO OIT\Asset Info')
writer = pd.ExcelWriter(os.path.join(savepath, 'Assets - CCOD ' + '-'.join(siteID_List) + '.xlsx'))
assetBook.to_excel(writer,'asset info')
writer.save()
writer.close()

# Envision
siteID_List = ['1003876149', '1003877862']
assetBook = getAssetInfo_forSites(siteID_List)
savepath = os.path.abspath(r'C:\Users\cohend\Documents\projects MDC\Envision')
writer = pd.ExcelWriter(os.path.join(savepath, 'Assets - OIT EFORT ' + '-'.join(siteID_List) + '.xlsx'))
assetBook.to_excel(writer,'asset info')
writer.save()
writer.close()


# Tarra
dunsList = ['941950321', '079364201', '072459266', '806345658', '007969868', '129673443', '793965997', '963023408', '132419532', '133898960', '786165220', '071012231', '002494128', '612283101', '063560986', '613447580', '074436841', '807419051', '062406718', '010634392', '177300043', '926941402', '619742203', '030437586', '837563717', '170472414', '017793860', '148247414', '121400063', '102333986', '041465720', '360709336', '033218871', '015302771', '122573988', '963519939', '006865273', '147283340', '127807704', '020115796', '062744495', '075769000', '043220623', '186917969', '007111818', '806345542', '150669716', '079764996', '056836364', '786528901', '808374677', '074457169', '878041102', '002629644', '079152306', '825470342', '086899924', '242388085', '064816847', '058072224', '135694847', '128599482', '063311922', '945580587', '003898285', '050245331']
assetBook = getAssetInfo_forDunsList(dunsList)
savepath = os.path.abspath(r'C:\Users\cohend\Documents\work data MDC\Administration\Commercial')
writer = pd.ExcelWriter(os.path.join(savepath, 'Assets - Tarra Blakstad ' + '.xlsx'))
assetBook.to_excel(writer,'asset info')
writer.save()
writer.close()

# Megan Bergman
dunsList = ['021192422', '015861677', '128248502', '831003012', '061724450', '174893891', '967189296', '623839003', '001056063', '102564572', '034392183', '010041247', '949301337', '966980773', '054929047', '617191122', '079603419', '174890590', '129106469', '050496447', '063322085', '145461500', '826284879', '650320716', '164570173', '123160249', '126930226', '807035527', '041957988', '179389650', '109067298', '069716736', '932442531', '006680487', '782205611', '956212542', '079128482', '018972739', '080269949', '031903446', '628009052', '055034086', '017290430', '084029057', '122032006', '004272808', '102100518', '102380953', '131274164', '796192243', '800314093']
assetBook = getAssetInfo_forDunsList(dunsList)
savepath = os.path.abspath(r'C:\Users\cohend\Documents\work data MDC\Administration\Commercial')
writer = pd.ExcelWriter(os.path.join(savepath, 'Assets - Megan Bergman 2 ' + '.xlsx'))
assetBook.to_excel(writer,'asset info')
writer.save()
writer.close()

#State of NM
dunsList = ['007111818']
assetBook = getAssetInfo_forDunsList(dunsList)
savepath = os.path.abspath(r'C:\Users\cohend\Documents\work data MDC\Administration\Commercial')
writer = pd.ExcelWriter(os.path.join(savepath, 'Assets - state of NM ' + '.xlsx'))
assetBook.to_excel(writer,'asset info')
writer.save()
writer.close()

#Pikes Peak Community College
SideIDList = ['10932273']
assetBook = getAssetInfo_forSites(SideIDList)
savepath = os.path.abspath(r'C:\Users\cohend\Documents\projects MDC\pikes peak community college')
writer = pd.ExcelWriter(os.path.join(savepath, 'Assets - PPCC' + '.xlsx'))
assetBook.to_excel(writer,'asset info')
writer.save()
writer.close()


quit()


asset_data = getAssetInfo_bySN(sn)

siteInfo_data = getSiteInfo_bySN(sn)

#datetime.strptime(asset_data['CONTRACT_SUBLINE_START_DATE'][0], '%Y-%m-%d %H:%M:%S').strftime("%B %d, %Y")
contractStartDate = parse(asset_data['CONTRACT_SUBLINE_START_DATE'][0])
contractStartDateText = contractStartDate.strftime("%B %d, %Y")

contractEndDate = parse(asset_data['CONTRACT_SUBLINE_END_DATE'][0])
contractEndDateText = contractEndDate.strftime("%B %d, %Y")

contractRemainingTime = relativedelta.relativedelta(contractEndDate, datetime.now())

yearsRemaining = contractRemainingTime.years
monthsRemaining = contractRemainingTime.months

print('Maintenance Contract\nStarted %s\nEnds %s\n%s years and %s months remaining as of %s' %(contractStartDateText,contractEndDateText,yearsRemaining,monthsRemaining,datetime.today().strftime('%m-%d-%y')))