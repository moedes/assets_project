import requests
import requests.exceptions
from requests_ntlm import HttpNtlmAuth
from requests.packages.urllib3.exceptions import InsecureRequestWarning
import pandas as pd
import credential
import os
import argparse
import json
import xlsxwriter

from datetime import datetime
from dateutil.parser import parse
from dateutil.relativedelta import relativedelta

requests.packages.urllib3.disable_warnings(InsecureRequestWarning)

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
   
   if asset_data['CONTRACT_SUBLINE_END_DATE'][0] != None:
      contractEndDate = parse(asset_data['CONTRACT_SUBLINE_END_DATE'][0])
      contractEndDateText = contractEndDate.strftime("%B %d, %Y")
      contractRemainingTime = relativedelta(contractEndDate, datetime.now())
      yearsRemaining = contractRemainingTime.years
      monthsRemaining = contractRemainingTime.months
      daysRemaining = contractRemainingTime.days
      maint_text = 'Maintenance Contract\nStarted %s\nEnds %s\n%s years and %s months and %s days remaining as of %s' %(contractStartDateText,contractEndDateText,yearsRemaining,monthsRemaining,daysRemaining,datetime.today().strftime('%m-%d-%y'))
      return maint_text
   else:
      return

def getAssetInfo_bySite(SiteID):

   session = requests.Session()
   session.auth = HttpNtlmAuth(credential.login['username'],credential.login['password'], session)
   r = session.get('http://sitelookup.corp.emc.com/models/getdataTLA.php?fld=Party%20ID&val=' + SiteID)
   
   print(json.dumps(r.text, indent=4))
   
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

   labels = ['CS_CUSTOMER_NAME', 'PARTY_NUMBER', 'PRODUCT_GROUP','MODEL','ITEM_SERIAL_NUMBER', 'CONTRACT_SUBLINE_END_DATE','ITEM_INSTALL_DATE','INSTALL_BASE_STATUS', 'CONNECT_IN_TYPE', 'CONNECT_HOME_TYPE', 'SYR_LAST_DIAL_HOME_DATE', 'CONTRACT_SUBLINE_STATUS', 'CONTRACT_SUBLINE_START_DATE','Address1', 'Address2', 'City', 'State', 'Province', 'Postal Code', 'Time Zone Name', 'DSM_EMAIL', 'DISTRICT', 'PRIMARY_CE_EMAIL']
   siteAssetInfo = inventory_df[labels]
   
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
   
   dunsAssetInfo = inventory_df

   return dunsAssetInfo

def getAssetInfo_forSites(siteID_List):
       
    for index, siteID in enumerate(siteID_List):
        
        if index == 0:
            assetBook = getAssetInfo_bySite(siteID)
        else:
            assetBook = pd.concat([assetBook, getAssetInfo_bySite(siteID)], ignore_index=True)  
    
    return assetBook

def getAssetInfo_forDunsList(dunsID_List):
       
    for index, dunsID in enumerate(dunsID_List):
        
        if index == 0:
            assetBook = getAssetInfo_byDuns(dunsID)
        else:
            assetBook = pd.concat([assetBook, getAssetInfo_byDuns(dunsID)], ignore_index=True)  
    
    labels = ['CS_CUSTOMER_NAME', 'PARTY_NUMBER', 'PRODUCT_GROUP','ITEM_SERIAL_NUMBER', 'CONTRACT_SUBLINE_END_DATE','ITEM_INSTALL_DATE','INSTALL_BASE_STATUS', 'CONNECT_IN_TYPE', 'CONNECT_HOME_TYPE', 'SYR_LAST_DIAL_HOME_DATE', 'CONTRACT_SUBLINE_STATUS', 'CONTRACT_SUBLINE_START_DATE','Address1', 'Address2', 'City', 'State', 'Province', 'Postal Code', 'Time Zone Name', 'DSM_EMAIL', 'DISTRICT', 'PRIMARY_CE_EMAIL']

    assetBook = assetBook[labels]

    return assetBook

def get_width(assetBook):
   widthdict = {}
   key = 0
   for label in assetBook:
      strlen = assetBook[label]
      maxlen = strlen.map(len, na_action='ignore').max()
      widthdict.update({key:maxlen})
      key += 1
   return widthdict

def sortren_book(assetBook):
   assetBook.rename(columns={"CS_CUSTOMER_NAME":"Customer Name","PARTY_NUMBER":"Site ID","PRODUCT_GROUP":"Product","MODEL":"Model","ITEM_SERIAL_NUMBER":"Serial Number"
                          , "ITEM_INSTALL_DATE": "Date Installed","CONTRACT_SUBLINE_END_DATE":"Contract End Date",
                          "INSTALL_BASE_STATUS":"Install Status", "CONNECT_IN_TYPE":"Dial In", "CONNECT_HOME_TYPE":"Dial Home", 
                          "SYR_LAST_DIAL_HOME_DATE":"Last Dial Home", "CONTRACT_SUBLINE_STATUS":"Contract Status", 
                          "CONTRACT_SUBLINE_START_DATE":"Contract Start Date", "DSM_EMAIL":"DSM email", 
                          "DISTRICT":"District", "PRIMARY_CE_EMAIL":"CE Email"}, inplace=True)
   assetBook['Contract End Date'] = pd.to_datetime(assetBook['Contract End Date'])
   date = datetime.today()
   date = date.today() - relativedelta(months=6)
   assetBook = assetBook[assetBook['Contract End Date'] > date]
   assetBook = assetBook.sort_values(by=['Contract End Date','Site ID'])

   return assetBook

def format_col(workbook, worksheet):
   formatcell = workbook.add_format({'align':'center'})

   for col, width in col_widths.items():
      width = width + 5
      worksheet.set_column(col,col,width, formatcell)
   
   return

def format_cells(workbook, worksheet, assetBook):
   date = datetime(2020,1,1)
   sixmonths = date.today() + relativedelta(months=6)
   twelvemonths = date.today() + relativedelta(months=12)
   today = date.today()
   red = workbook.add_format({'bg_color':'red','num_format': 'mmm d, yyyy'})
   orange = workbook.add_format({'bg_color':'#FF9933','num_format': 'mmm d, yyyy'})
   green = workbook.add_format({'bg_color':'green','num_format': 'mmm d, yyyy'})
   yellow = workbook.add_format({'bg_color':'#F7ED5F','num_format': 'mmm d, yyyy'})
   loc = assetBook.columns.get_loc("Contract End Date")
   length = len(assetBook)
   worksheet.conditional_format(1,loc,length,loc, {'type': 'date',
                                                 'criteria':'less than',
                                                 'value': today,
                                                 'format': red})
   worksheet.conditional_format(1,loc,length,loc, {'type': 'date',
                                                 'criteria':'greater than',
                                                 'value': twelvemonths,
                                                 'format': green})
   worksheet.conditional_format(1,loc,length,loc, {'type': 'date',
                                                 'criteria': 'between',
                                                 'minimum': today,
                                                 'maximum': sixmonths,
                                                 'format': orange})
   worksheet.conditional_format(1,loc,length,loc, {'type': 'date',
                                                 'criteria': 'between',
                                                 'minimum': sixmonths,
                                                 'maximum': twelvemonths,
                                                 'format': yellow})
   worksheet.freeze_panes(1,0)
   return

sn = ''
if sn != '':
   maintVisioText = getMaintenanceVisioText(sn)
   print(maintVisioText)

   siteID = getSiteID_bySN(sn)


   assetBook = getAssetInfo_bySite(siteID)
   savepath = os.path.abspath(r'C:\Users\mozesj\OneDrive - Dell Technologies\Documents\Python Projects\DellEMC\Assets\Janus')
   writer = pd.ExcelWriter(os.path.join(savepath, 'Assets - Site ID ' + siteID + '.xlsx'))
   assetBook.to_excel(writer,'asset info')
   writer.save()
   writer.close()

parser = argparse.ArgumentParser()
parser.add_argument("--cust", "-c", help="Valid Customer Input is Agilent, Janus, or SCL")

args = parser.parse_args()
if args.cust == "Janus":
   sites = ['2547453', '14735642','4257799','1003821648','63950','1003830326','1003868624','4285738']
elif args.cust == "SCL":
   sites = ['25992760','1003906813','2919727','14584171','10457835','4408824']
elif args.cust == "Agilent":
   sites = ['1004230935','4076980','4906201','1003828875','55363','1003831985','1003918035','14650376','3927620','12415254','1003828626','1003828315','3361376','1003828873','2367330','8957317','1003852457','1003828889','1003828292','2386956','1003828051','3654613','2939517','2419636','2419683','5275340','7960716','1003851155','1003855388','1003829005','4050166','81129','2113582','2419219','1003821912','1004541470','18696303','8067058','1003969723']
else:
   sites = ['2547453', '14735642','4257799','1003821648','63950','1003830326','1003868624','4285738']
   print("Use Valid Customer")

assetBook = getAssetInfo_forSites(sites)
col_widths = get_width(assetBook)
assetBook = sortren_book(assetBook)  

#for index, row in snassets.iterrows():
#   contractEndDate = parse(row['Contract End Date'])
#   contractEndDateText = contractEndDate.strftime("%B %d, %Y")
#   contractRemainingTime = relativedelta(contractEndDate, datetime.now())
#   yearsRemaining = contractRemainingTime.years
#   monthsRemaining = contractRemainingTime.months
#   daysRemaining = contractRemainingTime.days
#   maint_text = 'SN - %s Contract Ends %s %s years and %s months and %s days remaining as of %s' %(row['Serial Number'],contractEndDateText,yearsRemaining,monthsRemaining,daysRemaining,datetime.today().strftime('%m-%d-%y'))
#   print(maint_text)
#############################################

savepath = os.path.abspath(r'C:\Users\mozesj\OneDrive - Dell Technologies\Documents\Python Projects\DellEMC\Assets\Janus')
writer = pd.ExcelWriter(os.path.join(savepath, 'Assets - ' + args.cust + '.xlsx'), engine="xlsxwriter")
assetBook.to_excel(writer,'asset info', index=False)
workbook = writer.book
worksheet = writer.sheets['asset info']

format_col(workbook, worksheet)
format_cells(workbook, worksheet, assetBook)

writer.save()
writer.close()

exit()
