import xlwt
import xlrd
import pandas as pd
import schedule
import time
import os
import xlsxwriter
from xlutils.copy import copy
import sharepy as sharepy
import requests
import json

s = sharepy.connect("https://qonline.sharepoint.com",username="User@Organization.onmicrosoft.com", password="Password")
r = s.getfile("URL to the file", filename = 'file.extension')
print("Status Code : ",r.status_code)



revenueC = []
revListFinal = []
PlannedRevenue = []
ActualRevenue = []
complistFinal = []
PlannedCompensation = []
ActualCompensation = []
sublistFinal = []
PlannedSubContract = []
ActualSubContract = []
travellistFinal = []
PlannedTravel = []
ActualTravel = []
telelistFinal = []
PlannedTele = []
ActualTele = []
otherlistFinal = []
PlannedOther = []
ActualOther = []
alloclistFinal = []
PlannedAlloc = []
ActualAlloc = []
texpfinalList = []
PlannedTexp = []
ActualTexp = []
dmfinalList = []
PlannedDM = []
ActualDM = []
dmpfinalList = []
PlannedDMP = []
ActualDMP = []
offshorefinalList = []
PlannedOffshore = []
ActualOffshore = []
onsitefinalList = []
PlannedOnsite = []
ActualOnsite = []
bftefinalList = []
PlannedBFTE = []
ActualBFTE = []
a3fnallist =[]
PlannedA3 = []
ActualA3 = []
b1finallist = []
PlannedB1 = []
ActualB1 = []
b2finallist = []
PlannedB2 = []
ActualB2 = []
c1finallist = []
PlannedC1 = []
ActualC1 = []
c2finallist = []
PlannedC2 = []
ActualC2 = []
d1finallist = []
PlannedD1 = []
ActualD1= []
d2finallist = []
PlannedD2 = []
ActualD2 = []
d3finallist = []
PlannedD3 = []
ActualD3 = []
z1finallist = []
PlannedZ1 = []
ActualZ1 =[]
tftefinallist = []
PlannedTFTE =[]
ActualTFTE =[]
utilfinallist = []
PlannedUtil = []
ActualUtil =[]
ofbhfinallist = []
PlannedOFBH = []
ActualOFBH = []
onbhfinallist = []
PlannedONBH = []
ActualONBH = []
tbhfinallist = []
PlannedTBH = []
ActualTBH = []
shfinallist = []
PlannedSH = []
ActualSH = []
Revenue_list = []
result = []
month = []
SAPID = []
AccountList = []
filename = "test.xlsx"
# excel_file = xlwt.Workbook(filename)
# sheet = excel_file.add_sheet('SEP')
row = 0
col = 0
xf = 0
workbook = xlrd.open_workbook('input.xlsx')

sheet = workbook.sheet_by_index(0)
rowsno = sheet.nrows
colno = sheet.ncols


# -------------------------------Splitting Values of all compinies into different lists-----------------------------------
for A1P in range(sheet.ncols):
    i = 0
    while i <= sheet.ncols:
        try:
            A1P = [sheet.cell_value(rows, i) for rows in range(sheet.nrows)]
            for A1C in range(sheet.nrows):
                # cols = [sheet.cell_value(rows, 0) for rows in range(sheet.nrows)]
                A1C = [sheet.cell_value(rows, i) for rows in range(sheet.nrows)]
            for A1A in range(sheet.ncols):
                # cols = [sheet.cell_value(rows, 0) for rows in range(sheet.nrows)]
                A1A = [sheet.cell_value(rows, i + 1) for rows in range(sheet.nrows)]
                if i == sheet.ncols:
                    break
                else:
                    i = i + 1
                    C = A1A
                Revenue_list = A1A[4]
                SAPID.append(A1A[1])
                AccountList.append(A1A[2])
                revListFinal.append(Revenue_list)
                complistFinal.append(A1A[5])
                sublistFinal.append(A1A[6])
                travellistFinal.append(A1A[7])
                telelistFinal.append(A1A[8])
                otherlistFinal.append(A1A[9])
                alloclistFinal.append(A1A[10])
                texpfinalList.append(A1A[11])
                dmfinalList.append(A1A[12])
                dmpfinalList.append(A1A[13])
                offshorefinalList.append(A1A[57])
                onsitefinalList.append(A1A[58])
                bftefinalList.append(A1A[59])
                a3fnallist.append(A1A[45])
                b2finallist.append(A1A[46])
                b1finallist.append(A1A[47])
                c1finallist.append(A1A[48])
                c2finallist.append(A1A[49])
                d1finallist.append(A1A[50])
                d2finallist.append(A1A[51])
                d3finallist.append(A1A[52])
                z1finallist.append(A1A[53])
                tftefinallist.append(A1A[54])
                utilfinallist.append(A1A[64])
                ofbhfinallist.append(A1A[65])
                onbhfinallist.append(A1A[66])
                tbhfinallist.append(A1A[67])
                shfinallist.append(A1A[68])

                n = 0
        except IndexError:  # __________Catching index error of the list
            datais = 'null'
            break
            print("not working")
        break
    break
print("End of data!!!!")

# _________________________Splitting Actual and PLanned Values__________________________________
for i, v in enumerate(revListFinal):
    if i % 2 == 0:
        PlannedRevenue.append(v)
    else:
        ActualRevenue.append(v)
# PlannedRevenue = [round(x) for x in PlannedRevenue]
# ActualRevenue = [round(x) for x in ActualRevenue]

for i, v in enumerate(complistFinal):
    if i % 2 == 0:
        PlannedCompensation.append(v)
    else:
        ActualCompensation.append(v)
# PlannedCompensation = [round(x) for x in PlannedCompensation]
# ActualCompensation = [round(x) for x in ActualCompensation]

for i, v in enumerate(sublistFinal):
    if i % 2 == 0:
        PlannedSubContract.append(v)
    else:
        ActualSubContract.append(v)
# PlannedSubContract = [round(x) for x in PlannedSubContract]
# ActualSubContract = [round(x) for x in ActualSubContract]

for i, v in enumerate(travellistFinal):
    if i % 2 == 0:
        PlannedTravel.append(v)
    else:
        ActualTravel.append(v)
# PlannedTravel = [round(x) for x in PlannedTravel]
# ActualTravel = [round(x) for x in ActualTravel]

for i, v in enumerate(telelistFinal):
    if i % 2 == 0:
        PlannedTele.append(v)

    else:
        ActualTele.append(v)

# PlannedTele = [round(x) for x in PlannedTele]
# ActualTele = [round(x) for x in ActualTele]

for i, v in enumerate(otherlistFinal):
    if i % 2 == 0:
        PlannedOther.append(v)
    else:
        ActualOther.append(v)
# PlannedOther = [round(x) for x in PlannedOther]
# ActualOther = [round(x) for x in ActualOther]

for i, v in enumerate(alloclistFinal):
    if i % 2 == 0:
        PlannedAlloc.append(v)
    else:
        ActualAlloc.append(v)
# PlannedAlloc = [round(x) for x in PlannedAlloc]
# ActualAlloc = [round(x) for x in ActualAlloc]

for i, v in enumerate(texpfinalList):
    if i % 2 == 0:
        PlannedTexp.append(v)
    else:
        ActualTexp.append(v)
# PlannedTexp = [round(x) for x in PlannedTexp]
# ActualTexp = [round(x) for x in ActualTexp]

for i, v in enumerate(dmfinalList):
    if i % 2 == 0:
        PlannedDM.append(v)
    else:
        ActualDM.append(v)
# PlannedDM = [round(x) for x in PlannedDM]
# ActualDM = [round(x) for x in ActualDM]

for i, v in enumerate(dmpfinalList):
    if i % 2 == 0:
        PlannedDMP.append(v)
    else:
        ActualDMP.append(v)

PlannedDMP = [x * 100 for x in PlannedDMP]
ActualDMP = [x * 100 for x in ActualDMP]

# PlannedDMP = [round(x) for x in PlannedDMP]
# ActualDMP = [round(x) for x in ActualDMP]

for i, v in enumerate(offshorefinalList):
    if i % 2 == 0:
        PlannedOffshore.append(v)
    else:
        ActualOffshore.append(v)
# PlannedOffshore = [round(x) for x in PlannedOffshore]
# ActualOffshore = [round(x) for x in ActualOffshore]

for i, v in enumerate(onsitefinalList):
    if i % 2 == 0:
        PlannedOnsite.append(v)
    else:
        ActualOnsite.append(v)
# PlannedOnsite = [round(x) for x in PlannedOnsite]
# ActualOnsite = [round(x) for x in ActualOnsite]

for i, v in enumerate(bftefinalList):
    if i % 2 == 0:
        PlannedBFTE.append(v)
    else:
        ActualBFTE.append(v)
# PlannedBFTE = [round(x) for x in PlannedBFTE]
# ActualBFTE = [round(x) for x in ActualBFTE]

for i, v in enumerate(a3fnallist):
    if i % 2 == 0:
        PlannedA3.append(v)
    else:
        ActualA3.append(v)
# PlannedA3 = [round(x) for x in PlannedA3]
# ActualA3 = [round(x) for x in ActualA3]

for i, v in enumerate(b1finallist):
    if i % 2 == 0:
        PlannedB1.append(v)
    else:
        ActualB1.append(v)
# PlannedB1 = [round(x) for x in PlannedB1]
# ActualB1 = [round(x) for x in ActualB1]

for i, v in enumerate(b2finallist):
    if i % 2 == 0:
        PlannedB2.append(v)
    else:
        ActualB2.append(v)
# PlannedB2 = [round(x) for x in PlannedB2]
# ActualB2 = [round(x) for x in ActualB2]

for i, v in enumerate(c1finallist):
    if i % 2 == 0:
        PlannedC1.append(v)
    else:
        ActualC1.append(v)
# PlannedC1 = [round(x) for x in PlannedC1]
# ActualC1 = [round(x) for x in ActualC1]

for i, v in enumerate(c2finallist):
    if i % 2 == 0:
        PlannedC2.append(v)
    else:
        ActualC2.append(v)
# PlannedC2 = [round(x) for x in PlannedC2]
# ActualC2 = [round(x) for x in ActualC2]

for i, v in enumerate(d1finallist):
    if i % 2 == 0:
        PlannedD1.append(v)
    else:
        ActualD1.append(v)
# PlannedD1 = [round(x) for x in PlannedD1]
# ActualD1 = [round(x) for x in ActualD1]

for i, v in enumerate(d2finallist):
    if i % 2 == 0:
        PlannedD2.append(v)
    else:
        ActualD2.append(v)

# PlannedD2 = [round(x) for x in PlannedD2]
# ActualD2 = [round(x) for x in ActualD2]

for i, v in enumerate(d3finallist):
    if i % 2 == 0:
        PlannedD3.append(v)
    else:
        ActualD3.append(v)
# PlannedD3 = [round(x) for x in PlannedD3]
# ActualD3 = [round(x) for x in ActualD3]

for i, v in enumerate(z1finallist):
    if i % 2 == 0:
        PlannedZ1.append(v)
    else:
        ActualZ1.append(v)
# PlannedZ1 = [round(x) for x in PlannedZ1]
# ActualZ1 = [round(x) for x in ActualZ1]

for i, v in enumerate(tftefinallist):
    if i % 2 == 0:
        PlannedTFTE.append(v)
    else:
        ActualTFTE.append(v)
# PlannedTFTE = [round(x) for x in PlannedTFTE]
# ActualTFTE = [round(x) for x in ActualTFTE]

for i, v in enumerate(utilfinallist):  ###########################-------------------Percentage field---------------------------------------
    if i % 2 == 0:
        PlannedUtil.append(v)
    else:
        ActualUtil.append(v)
PlannedUtil = [x * 100 for x in PlannedUtil]
ActualUtil = [x * 100 for x in ActualUtil]
# PlannedUtil = [round(x) for x in PlannedUtil]
# ActualUtil = [round(x) for x in ActualUtil]

for i, v in enumerate(ofbhfinallist):
    if i % 2 == 0:
        PlannedOFBH.append(v)
    else:
        ActualOFBH.append(v)
# PlannedOFBH = [round(x) for x in PlannedOFBH]
# ActualOFBH = [round(x) for x in ActualOFBH]

for i, v in enumerate(onbhfinallist):
    if i % 2 == 0:
        PlannedONBH.append(v)
    else:
        ActualONBH.append(v)
# PlannedONBH = [round(x) for x in PlannedONBH]
# ActualONBH = [round(x) for x in ActualONBH]

for i, v in enumerate(tbhfinallist):
    if i % 2 == 0:
        PlannedTBH.append(v)
    else:
        ActualTBH.append(v)
# PlannedTBH = [round(x) for x in PlannedTBH]
# ActualTBH = [round(x) for x in ActualTBH]

for i, v in enumerate(shfinallist):
    if i % 2 == 0:
        PlannedSH.append(v)
    else:
        ActualSH.append(v)
# PlannedSH = [round(x) for x in PlannedSH]
# ActualSH = [round(x) for x in ActualSH]



# ___________________________Removing NULL values from column Names_______________________________
while "" in A1P:
    A1P.remove("")
while "" in AccountList:
    AccountList.remove("")
while "" in SAPID:
    SAPID.remove("")

# for i in AccountList:
#     print(i)
L = []
for i in A1P:
    A = i
    P = i
    L.append(P)
    L.append(A)
L[0] = "Account"

# print(L[2])

pd.DataFrame(L).to_excel('TopCols.xlsx', header=False, index=False)
sheetsy = pd.read_excel('TopCols.xlsx', sheet_name=None)
# all sheets

df = pd.DataFrame.from_dict(
    {'month': L[2],'accountID':SAPID,'accountName': AccountList,'plannedReveneu': PlannedRevenue, 'actualRevenue': ActualRevenue,
     'plannedEmployeeCost': PlannedCompensation, 'actualEmployeeCost': ActualCompensation,
     'plannedContractorCost': PlannedSubContract, 'actualContractorcost': ActualSubContract, 'plannedTravelCost': PlannedTravel,
     'actualTravelCost': ActualTravel,
     'plannedTelecomCost': PlannedTele, 'actualTelecomCost': ActualTele, 'plannedOtherCost': PlannedOther, 'actualOtherCost': ActualOther,
     'plannedAllocCost': PlannedAlloc, 'actualAllocCost': ActualAlloc,
     'plannedTotalExpense': PlannedTexp, 'actualTotalExpense': ActualTexp, 'plannedMargin': PlannedDM, 'actualMargin': ActualDM,
     'plannedA3': PlannedA3, 'actualA3': ActualA3, 'plannedB1': PlannedB1,
     'actualB1': ActualB1, 'plannedB2': PlannedB2, 'actualB2': ActualB2,
     'plannedC1': PlannedC1, 'aCtualC1': ActualC1, 'plannedC2': PlannedC2, 'actualC2': ActualC2, 'plannedD1': PlannedD1, 'actualD1': ActualD1,
     'plannedD2': PlannedD2, 'actualD2': ActualD2,
     'plannedD3': PlannedD3, 'actualD3': ActualD3, 'plannedZ1': PlannedZ1, 'actualZ1': ActualZ1, 'plannedTotalFTE': PlannedTFTE,
     'actualTotalFTE': ActualTFTE, 'plannedBillingOffshore': PlannedOFBH, 'actualBillingOffshore': ActualOFBH, 'plannedBillingOnsite': PlannedONBH,
     'actualBillingOnsite': ActualONBH, 'plannedTotalBilling': PlannedTBH, 'plannedBilledUtil': PlannedUtil, 'actualBilledUtil': ActualUtil,
     'actualTotalBilling': ActualTBH, 'P.StandardHrs': PlannedSH, 'A.StandardHrs': ActualSH})

print('----------------------------------------------------------')
# print(df)
# rowindex = 2

df.to_excel('final.xlsx', "test", index=False, header=True)

excel_data_df = pd.read_excel("final.xlsx")

json_str = excel_data_df.to_json(orient='records')



auth_token='eyJhbGciOiJIUzUxMiJ9.eyJzdWIiOiJzaGFyYXRoYSIsImF1dGgiOiJST0xFX1VTRVIiLCJleHAiOjE1NzU2MjI2ODh9.EvmC5TxNGvgXf9kinPCX8qfUZBCtbKrsuYcUZFZ41b3jUKIdU0_n8zuTOtlzlgkLsI9WpjwdiHZtgli3UqugXw'
hed = {'Authorization': 'Bearer ' + auth_token}
jsonstringconv = json.loads(json_str)
data = jsonstringconv
print('Excel Sheet to JSON:\n', jsonstringconv)
url = 'https://ddpowerbi-test.quinnox.info/api/account-financedetailsList'
response = requests.post(url, json=data, headers=hed)
print(response)
print(response.content)
