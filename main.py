import pandas as pd
import numpy as np
import openpyxl

#Reads the Adobe financial calendar
financialcalendar = r"\\sjshare\finance\Treasury\Financial_Database\ADOBE_FINANCIAL_CALENDAR.xlsx"
sheet = "ADBE_cal"
df_cal = pd.read_excel(financialcalendar, sheet_name = sheet)

#Based on today's date, displays all periods that end after or on today's date
print(df_cal.head(20))
print(df_cal.columns)
print(df_cal["Quarter"])
print(df_cal["Quarter"].value_counts())
today = pd.Timestamp.today()
print(today)
df_cal = df_cal[df_cal["Per_End"]>= today]
print(df_cal.head(10))

#Selects only the Year, Quarter, Qtr_Ticker, Qtr_Start, and Qtr_End
df_cal.info()
df_cal = df_cal[["Year", "Quarter", "Qtr_Ticker", "Qtr_Start", "Qtr_End"]]
print(df_cal.head(10))

#Because our forecast is quarterly, from current quarter and on out, there could be some duplicate quarter info, so the duplicates here are removed
df_cal = df_cal.drop_duplicates()
print(df_cal.head(10))

#Reads in the Equity Comp database
equitydatabase = r"\\sjshare\finance\Treasury\Financial_Database\EQUITY_COMP_FULL.xlsm"
sheet2 = "RSU"
df_rsu = pd.read_excel(equitydatabase, sheet_name = sheet2)
print(df_rsu.head(10))

#In reality, shares withheld are whole shares, so the SHS W/H is rounded up to the nearest number for each row
df_rsu["SHS W/H"] = df_rsu["SHS W/H"].apply(np.ceil)
print(df_rsu.head(10))

#Subtracts rounded-up SHS W/H from shares vested to get vested SHS post-withholding
df_rsu["Column1"] = df_rsu["Vest_Shares"]-df_rsu["SHS W/H"]
print(df_rsu.head(10))
df_rsu = df_rsu.rename(columns={"Column1":"Vested SHS Post-W/H"})
print(df_rsu.head(10))

#Selects grant number and grant date from the previous dataframe, drops duplicates to get a unique RSU grant with its own unique grant date
grantdates = df_rsu[["Grant_Number","Grant_Date"]]
grantdates = grantdates.drop_duplicates()
print(grantdates.head(10))

#Creates another dataframe which will show the maximum vest date for each grant number (or the final vest date)
groupedgrants = df_rsu.groupby(["Grant_Number"])["Vest_Date"].max()
pd.set_option("display.max_rows", None,"display.max_columns", None)
print(groupedgrants)
print(type(grantdates))
print(type(groupedgrants))

#Combines the two previous dataframes (grantdates, groupedgrants), calculates Total Expense Days by subtracting each grant date from each final vest date
rsudates = grantdates.merge(groupedgrants, on="Grant_Number")
rsudates["Total Expense Days"] = ((rsudates["Vest_Date"]-rsudates["Grant_Date"]).dt.days)+1
print(rsudates.head(10))

#Creates another dataframe to show the total number of shares that will vest for each grant number
grantedshares = df_rsu.groupby(["Grant_Number"])["Vest_Shares"].sum()
print(type(grantedshares))

#Creates another dataframe to pick Grant Number and Grant Price, drops duplicates to get 1 unique grant price for 1 unique grant number. Combines grantedshares and this dataframe.
grantprice = df_rsu[["Grant_Number", "Grant_Price"]]
grantprice = grantprice.drop_duplicates()
grantinfo = grantprice.merge(grantedshares, on="Grant_Number")
print(grantinfo.head(10))
grantinfo = grantinfo.rename(columns={"Vest_Shares":"Shares Granted"})

#Calculates total expense by multiplying the grant price by the total number of shares that will vest for each grant number (aka shares granted), assigns the field to rsudates dataframe, and sets rsuinfo to rsudates dataframe
grantinfo["Total Expense"] = grantinfo["Grant_Price"]*grantinfo["Shares Granted"]
print(grantinfo.head(10))
rsudates["Total Expense"] = grantinfo["Total Expense"]
rsuinfo = rsudates
print(rsuinfo.head(10))

#Gets the number of rows in rsuinfo dataframe (the number of unique RSU grant numbers) and the number of quarter-ends given today's date
index = rsuinfo.index
rows = len(index)
print(rows)
index2 = df_cal["Qtr_End"].index
rows2 = len(index2)
print(rows2)
print(rsuinfo.values)
print(type(rsuinfo.values))

#With another dataframe, repeats each RSU grant number's data the number of quarter-ends given today's date
replicatedrsuinfo = pd.DataFrame(np.repeat(rsuinfo.values,rows2,axis=0))
replicatedrsuinfo.columns = rsuinfo.columns
print(replicatedrsuinfo.head(82))
print(type(df_cal["Qtr_End"]))
print(df_cal.dtypes)
print(replicatedrsuinfo.dtypes)

#With another dataframe, repeats the series of quarter-ends by the number of unique RSU grant numbers and adds that as a new column in replicatedrsuinfo
replicatedqtrend = pd.concat([df_cal["Qtr_End"]]*rows, ignore_index=True)
print(replicatedqtrend.head(82))
print(type(replicatedqtrend))
replicatedrsuinfo["Qtr_End"] = replicatedqtrend
replicatedrsuinfo = replicatedrsuinfo.rename(columns={"Vest_Date":"Final Vest Date"})

#Creates another dataframe, amorttable, equal to replicatedrsuinfo, and calculates Expense Days by subtracting each grant date from each quarter-end date. A new column takes the minimum of Total Expense Days in each row and Expense Days just calculated, drops the earlier Expense Days calculation.
amorttable = replicatedrsuinfo
print(amorttable.head(82))
amorttable["Expense_Days"] = ((amorttable["Qtr_End"] - amorttable["Grant_Date"]).dt.days)+1
amorttable["Expense Days"] = amorttable[["Total Expense Days", "Expense_Days"]].min(axis=1)
amorttable = amorttable.drop(["Expense_Days"],axis=1)
print(amorttable.head(82))

#Calculates amortized expense by multiplying each total expense by the number of expense days divided by total expense days.
amorttable["Amortized Expense"] = amorttable["Total Expense"]*(amorttable["Expense Days"]/amorttable["Total Expense Days"])
print(amorttable.head(82))

#Creates another dataframe repeating the series of quarter start dates by the number of unique RSU grant numbers and adds that as a new column to amorttable.
replicatedqtrstart = pd.concat([df_cal["Qtr_Start"]]*rows, ignore_index=True)
amorttable["Qtr_Start"] = replicatedqtrstart
print(amorttable.head(82))

#Creates a subset of the amorttable containing just the current quarter's start date.
df_cal2 = df_cal.reset_index(drop=True)
print(df_cal2.head(10))
subamorttable = amorttable[amorttable["Qtr_Start"]== df_cal2.at[0,"Qtr_Start"]]
print(subamorttable.head(10))

#Creates a sub-subset containing grant dates that are greater than or equal to quarter-start dates.
subamorttable2 = subamorttable[subamorttable["Grant_Date"]>= subamorttable["Qtr_Start"]]

#Renames Total Expense to BOQ Unamortized Expense, and calculates EOQ Unamortized Expense by subtracting Amortized Expense from Total Expense
subamorttable2["BOQ Unamortized Expense"] = subamorttable2["Total Expense"]
subamorttable2["EOQ Unamortized Expense"] = subamorttable2["Total Expense"] - subamorttable2["Amortized Expense"]

#Calculates Days In Quarter by subtracting the quarter start date from the quarter-end date. Then, calculates Average Unamortized Expense by taking the average of BOQ and EOQ Unamortized Expense, and multiplying it by Expense Days over Days In Quarter.
subamorttable2["Days In Quarter"] = ((subamorttable2["Qtr_End"] - subamorttable2["Qtr_Start"]).dt.days)+1
subamorttable2["Average Unamortized Expense"] = ((subamorttable2["BOQ Unamortized Expense"] + subamorttable2["EOQ Unamortized Expense"])/(2))*((subamorttable2["Expense Days"])/(subamorttable2["Days In Quarter"]))
print(subamorttable2.head(10))

#Creates another sub-subset containing grant dates that are less than the quarter-start dates. Calculates Prev Quarter End by subtracting each quarter-start date by 1 day.
subamorttable3 = subamorttable[subamorttable["Grant_Date"]< subamorttable["Qtr_Start"]]
subamorttable3["Prev Qtr_End"] = subamorttable3["Qtr_Start"]-pd.Timedelta(1,unit='d')

#Calculates previous quarter expense days by subtracting grant date from the previous quarter-end date. New column takes the minimum of total expense days and previous quarter expense days, drops the previous quarter expense days calculation.
subamorttable3["Prev Qtr Expense_Days"] = ((subamorttable3["Prev Qtr_End"] - subamorttable3["Grant_Date"]).dt.days)+1
subamorttable3["Prev Qtr Expense Days"] = subamorttable3[["Total Expense Days", "Prev Qtr Expense_Days"]].min(axis=1)
subamorttable3 = subamorttable3.drop(["Prev Qtr Expense_Days"],axis=1)

#Calculates previous quarter amortized expense by multiplying total expense by previous quarter expense days divided by total expense days. Calculated Previous Quarter EOQ Unamortized Expense by subtracting previous quarter amortized expense by total expense.
subamorttable3["Prev Qtr Amortized Expense"] = subamorttable3["Total Expense"]*(subamorttable3["Prev Qtr Expense Days"]/subamorttable3["Total Expense Days"])
subamorttable3["Prev Qtr EOQ Unamortized Expense"] = subamorttable3["Total Expense"]-subamorttable3["Prev Qtr Amortized Expense"]

#The BOQ Unamortized Expense equals the previous quarter EOQ Unamortized Expense, and EOQ Unamortized Expense is equal to the total expense minus the amortized expense. The Average Unamortized Expense is just an average of BOQ and EOQ Unamortized Expense.
subamorttable3["BOQ Unamortized Expense"] = subamorttable3["Prev Qtr EOQ Unamortized Expense"]
subamorttable3["EOQ Unamortized Expense"] = subamorttable3["Total Expense"] - subamorttable3["Amortized Expense"]
subamorttable3["Average Unamortized Expense"] = ((subamorttable3["BOQ Unamortized Expense"] + subamorttable3["EOQ Unamortized Expense"])/(2))
print(subamorttable3.head(10))

#Creates a subset of amorttable where the quarter start dates do not equal the current quarter's start date. Calculates Prev Quarter End by subtracting each quarter-start date by 1 day.
subamorttable4 = amorttable[amorttable["Qtr_Start"]!= df_cal2.at[0,"Qtr_Start"]]
subamorttable4["Prev Qtr_End"] = subamorttable4["Qtr_Start"]-pd.Timedelta(1,unit='d')

#Calculates previous quarter expense days by subtracting grant date from the previous quarter-end date. New column takes the minimum of total expense days and previous quarter expense days, drops the previous quarter expense days calculation.
subamorttable4["Prev Qtr Expense_Days"] = ((subamorttable4["Prev Qtr_End"] - subamorttable4["Grant_Date"]).dt.days)+1
subamorttable4["Prev Qtr Expense Days"] = subamorttable4[["Total Expense Days", "Prev Qtr Expense_Days"]].min(axis=1)
subamorttable4 = subamorttable4.drop(["Prev Qtr Expense_Days"],axis=1)

#Calculates previous quarter amortized expense by multiplying total expense by previous quarter expense days divided by total expense days. Calculated Previous Quarter EOQ Unamortized Expense by subtracting previous quarter amortized expense by total expense.
subamorttable4["Prev Qtr Amortized Expense"] = subamorttable4["Total Expense"]*(subamorttable4["Prev Qtr Expense Days"]/subamorttable4["Total Expense Days"])
subamorttable4["Prev Qtr EOQ Unamortized Expense"] = subamorttable4["Total Expense"]-subamorttable4["Prev Qtr Amortized Expense"]

#The BOQ Unamortized Expense equals the previous quarter EOQ Unamortized Expense, and EOQ Unamortized Expense is equal to the total expense minus the amortized expense. The Average Unamortized Expense is just an average of BOQ and EOQ Unamortized Expense.
subamorttable4["BOQ Unamortized Expense"] = subamorttable4["Prev Qtr EOQ Unamortized Expense"]
subamorttable4["EOQ Unamortized Expense"] = subamorttable4["Total Expense"] - subamorttable4["Amortized Expense"]
subamorttable4["Average Unamortized Expense"] = ((subamorttable4["BOQ Unamortized Expense"] + subamorttable4["EOQ Unamortized Expense"])/(2))
print(subamorttable4.head(82))

#Ensures that the three subsets of the amorttable in some form or the other each have the same number of columns and column names in the same order.
subamorttable2 = subamorttable2[["Grant_Number","Grant_Date","Final Vest Date","Total Expense Days","Total Expense","Qtr_End","Expense Days","Amortized Expense","Qtr_Start","BOQ Unamortized Expense","EOQ Unamortized Expense","Average Unamortized Expense"]]
subamorttable3 = subamorttable3[["Grant_Number","Grant_Date","Final Vest Date","Total Expense Days","Total Expense","Qtr_End","Expense Days","Amortized Expense","Qtr_Start","BOQ Unamortized Expense","EOQ Unamortized Expense","Average Unamortized Expense"]]
subamorttable4 = subamorttable4[["Grant_Number","Grant_Date","Final Vest Date","Total Expense Days","Total Expense","Qtr_End","Expense Days","Amortized Expense","Qtr_Start","BOQ Unamortized Expense","EOQ Unamortized Expense","Average Unamortized Expense"]]

#Combines the three subsets, this is the complete, new amorttable. Amorttable is sorted by grant number and quarter-end dates such that the grant numbers are in ascending order with the quarter-end dates for each grant number in ascending order as well.
amorttable = pd.concat([subamorttable2,subamorttable3,subamorttable4],axis=0,ignore_index=True)
amorttable = amorttable.sort_values(by=["Grant_Number","Qtr_End"],ignore_index=True)
print(amorttable.head(82))

#Creates a subset of amorttable where the quarter end dates are equal to the current quarter quarter-end date.
subamorttable5 = amorttable[amorttable["Qtr_End"]== df_cal2.at[0,"Qtr_End"]]
subamorttable5 = subamorttable5.reset_index(drop=True)
print(subamorttable5.head(82))

#Reads ADBE Historicals, sheet "CurrentQAvgSharePrice" in. IMPORTANT: Before this entire code is executed, this file should be refreshed to get the latest current quarter average share price estimate.
currentqavgshareprice = r"\\sjshare\finance\Treasury\Financial_Database\ADBE Historicals.xlsm"
sheetininterest = "CurrentQAvgSharePrice"
cqavgshareprice = pd.read_excel(currentqavgshareprice, sheet_name = sheetininterest)

#Equates dataframe cqavgshareprice to just column Current Quarter Est Avg Share Price, so it's now a Series.
cqavgshareprice = cqavgshareprice["Current Quarter Est Avg Share Price"]
print(cqavgshareprice.head(3))

#Equates cqavgshareprice again to just the actual current quarter est avg share price number (one value), and makes that into a Series.
cqavgshareprice = cqavgshareprice.iloc[0]
print(cqavgshareprice)
cqavgshareprice = pd.Series(cqavgshareprice)
print(cqavgshareprice)

#Gets the number of rows from subset of amorttable (subamorttable5). This will be the number of unique grant numbers. Repeats cqavgshareprice Series the number of unique grant numbers times.
index3 = subamorttable5.index
rows3 = len(index3)
print(rows3)
cqavgshareprice = cqavgshareprice.repeat(rows3)
cqavgshareprice = cqavgshareprice.reset_index(drop=True)

#Combines cqavgshareprice Series with subamorttable5
subamorttable5 = pd.concat([subamorttable5,cqavgshareprice],axis=1)
subamorttable5 = subamorttable5.rename(columns={0:"Current Quarter Est Avg Share Price"})
print(subamorttable5.head(82))

#Creates another subset of amorttable where quarter ends do not equal the current quarter quarter-end.
subamorttable6 = amorttable[amorttable["Qtr_End"]!= df_cal2.at[0,"Qtr_End"]]
print(subamorttable6.head(82))

#Adds a column called Est Quarter Avg Share Price to dataframe containing list of relevant forecasted quarter-ends.
df_cal3 = df_cal2
df_cal3["Est Quarter Avg Share Price"] = pd.Series(dtype="float64")
print(df_cal3)

#Reads in ADBE Historicals file again, ultimately makes cqavgshareprice2 equivalent to the value of the current quarter est avg share price.
cqavgshareprice2 = pd.read_excel(currentqavgshareprice, sheet_name = sheetininterest)
cqavgshareprice2 = cqavgshareprice2["Current Quarter Est Avg Share Price"]
cqavgshareprice2 = cqavgshareprice2.iloc[0]

#In the first row of df_cal3 of column "Est Quarter Avg Share Price", sets this equal to cqavgshareprice2.
df_cal3.at[0,"Est Quarter Avg Share Price"] = cqavgshareprice2
print(df_cal3)

#Creates a loop, calculates every future quarter's avg share price by taking the previous quarter's avg share price and multiplying it by (1.1)^0.25.
for i in range(1,len(df_cal3)):
    df_cal3.loc[i,"Est Quarter Avg Share Price"]=df_cal3.loc[i-1,"Est Quarter Avg Share Price"]*((1.1)**(0.25))
print(df_cal3)

#Equates fqavgshareprice Series to Est Quarter Avg Share Price column just calculated, and drops the first row (since we now just want the future quarter's avg share prices).
fqavgshareprice = df_cal3["Est Quarter Avg Share Price"]
fqavgshareprice = fqavgshareprice.drop(labels=[0])
fqavgshareprice = fqavgshareprice.reset_index(drop=True)
subamorttable6 = subamorttable6.reset_index(drop=True)

#Creates another dataframe which includes just the number of unique grant numbers from subamorttable6, and counts its number of rows.
subamorttable6uniquegrant = subamorttable6.drop_duplicates(subset="Grant_Number")
print(subamorttable6uniquegrant.head(82))
index4 = subamorttable6uniquegrant.index
rows4 = len(index4)
print(rows4)

#Creates replicatedfqavgshareprice Series where the future quarter's avg share prices are repeated by the number of unique grant numbers.
replicatedfqavgshareprice = pd.concat([fqavgshareprice]*rows4, ignore_index=True)
print(replicatedfqavgshareprice.head(150))

#Combines subamorttable6 and replicatedfqavgshareprice, renames subamorttable5 and subamorttable6 avg share price columns so it says the same label.
subamorttable6 = pd.concat([subamorttable6,replicatedfqavgshareprice],axis=1)
subamorttable5 = subamorttable5.rename(columns={"Current Quarter Est Avg Share Price":"Est Avg Share Price In Quarter"})
subamorttable6 = subamorttable6.rename(columns={"Est Quarter Avg Share Price":"Est Avg Share Price In Quarter"})
print(subamorttable6.head(150))

#Combines subamorttable5 and subamorttable6 to get the new, updated amorttable. Again, Amorttable is sorted by grant number and quarter-end dates such that the grant numbers are in ascending order with the quarter-end dates for each grant number in ascending order as well.
amorttable = pd.concat([subamorttable5,subamorttable6],axis=0,ignore_index=True)
amorttable = amorttable.sort_values(by=["Grant_Number","Qtr_End"],ignore_index=True)

#Calculates Buy Back Shares by dividing Average Unamortized Expense by Est Avg Share Price In Quarter
amorttable["Buy Back Shares"]=amorttable["Average Unamortized Expense"]/amorttable["Est Avg Share Price In Quarter"]
print(amorttable.head(82))
print(df_rsu.head(82))

#Using existing dataframe df_rsu, creates df_rsu2 which shows the number of shares vesting post-withholding by each vesting date. Ensures that these vest dates are greater than or equal to the current quarter's quarter start date.
df_rsu2 = df_rsu.groupby(["Vest_Date"],as_index=False)["Vested SHS Post-W/H"].sum()
df_rsu2 = df_rsu2[df_rsu2["Vest_Date"]>= df_cal3.at[0,"Qtr_Start"]]

#Creates a copy of dataframe (df_cal4) containing all relevant quarter-end dates, and tacks on an empty column "Shares Outstanding".
df_cal4 = df_cal.reset_index(drop=True)
df_cal4["Shares Outstanding"] = pd.Series(dtype="float64")
print(df_cal4)

#Creates another dataframe, sharesoutstanding, equal to subset of above dataframe showing just the quarter-end dates and shares outstanding columns.
sharesoutstanding = df_cal4[["Qtr_End","Shares Outstanding"]]
print(sharesoutstanding)
print(df_rsu2)

#Calculates Shares Outstanding by getting the total number of shares post-withholding vested at dates less than or equal to each quarter-end dates in Shares Outstanding. Utilizes df_rsu2 containing vesting schedule.
sharesoutstanding["Shares Outstanding"]=sharesoutstanding.apply(lambda x: df_rsu2.loc[(df_rsu2.Vest_Date<=x.Qtr_End),"Vested SHS Post-W/H"].sum(),axis=1)

#Creates another column Qtr_Start in sharesoutstanding equal to df_cal4's Qtr_Start column. Creates BOQ Shares Outstanding blank column in sharesoutstanding.
sharesoutstanding["Qtr_Start"]=df_cal4["Qtr_Start"]
sharesoutstanding["BOQ Shares Outstanding"]=pd.Series(dtype="float64")

#Calculates BOQ Shares Outstanding by getting the total number of shares post-withholding vested at dates less than or equal to each quarter-start date in Shares Outstanding. Utilizes df_rsu2 containing vesting schedule.
sharesoutstanding["BOQ Shares Outstanding"]=sharesoutstanding.apply(lambda x: df_rsu2.loc[(df_rsu2.Vest_Date<=x.Qtr_Start),"Vested SHS Post-W/H"].sum(),axis=1)
print(sharesoutstanding)

#For each vest date in df_rsu2, creates a loop so that it provides the quarter-end of the quarter in which the vest date appears. This appears in a list. Concatenates each calculated quarter-end date into a series and adds it as new column Qtr_End in df_rsu2
qtr_end = []
for date in df_rsu2["Vest_Date"]:
    qtr_end.append((sharesoutstanding.loc[(sharesoutstanding["Qtr_Start"]<=date)&(sharesoutstanding["Qtr_End"]>=date),"Qtr_End"]))
qtr_end = pd.concat(qtr_end,ignore_index=True)
df_rsu2["Qtr_End"]=qtr_end

#In df_rsu2, calculates "Relevant Days" by subtracting each vest date from its quarter-end date.
df_rsu2["Relevant Days"]= ((df_rsu2["Qtr_End"] - df_rsu2["Vest_Date"]).dt.days)+1

#Creates a copy of dataframe (df_cal5) containing all relevant dates in this forecast, and chooses only its Qtr Start and Qtr End contents.
df_cal5 = df_cal.reset_index(drop=True)
df_cal5 = df_cal5[["Qtr_Start","Qtr_End"]]
print(df_cal5)

#Combines df_rsu2 and df_cal5, calculates Days In Quarter for df_rsu2 by subtracting the quarter start date from the quarter end date.
df_rsu2 = pd.merge(df_rsu2,df_cal5,on="Qtr_End",how="inner")
df_rsu2["Days In Quarter"]=((df_rsu2["Qtr_End"]-df_rsu2["Qtr_Start"]).dt.days)+1

#Calculates Qtr Weight by multiplying Vested SHS Post-W/H by Relevant Days over Days In Quarter. Creates df_rsu3 dataframe which sums the quarter weight column by quarter-end dates.
df_rsu2["Qtr Weight"]=df_rsu2["Vested SHS Post-W/H"]*((df_rsu2["Relevant Days"])/(df_rsu2["Days In Quarter"]))
print(df_rsu2)
df_rsu3 = df_rsu2.groupby(["Qtr_End"],as_index=False)["Qtr Weight"].sum()
print(df_rsu3)

#Combines df_rsu3 with sharesoutstanding and fills any N/A values in Qtr Weight with 0.
sharesoutstanding = pd.merge(sharesoutstanding,df_rsu3,on="Qtr_End",how="left")
sharesoutstanding["Qtr Weight"]=sharesoutstanding["Qtr Weight"].fillna(0)
print(sharesoutstanding)

#Calculates Total Weighted Average by adding BOQ Shares Outstanding and Qtr Weight together. Rounds all float numbers in this sharesoutstanding dataframe to 6 decimal places.
sharesoutstanding["Total Weighted Average"]=sharesoutstanding["BOQ Shares Outstanding"]+sharesoutstanding["Qtr Weight"]
print(sharesoutstanding)
pd.options.display.float_format="{:.6f}".format
print(sharesoutstanding)

#From earlier dataframe, grantinfo, creates variable totalsharesgranted equivalent to the sum of all shares granted on file.
totalsharesgranted=grantinfo["Shares Granted"].sum()
print(totalsharesgranted)

#Creates a copy of dataframe (df_cal6) containing all relevant dates in this forecast, and creates column Total Shares Granted equivalent to variable totalsharesgranted (which is copied down all this dataframe's rows).
df_cal6 = df_cal.reset_index(drop=True)
print(df_cal6)
df_cal6["Total Shares Granted"]=totalsharesgranted

#Creates a copy of dataframe df_rsu called df_rsu4 and calculates the sum of all vested shares by vest date. Then, creates a new column (BOQ Shares Outstanding Without Withholding) in df_cal6 where each row is equal to the total number of shares vested at dates less than or equal to each quarter-start date. Utilizes df_rsu4 containing vesting schedule.
#Finally, creates another column in df_cal6 (Unvested Shares BOQ) by subtracting BOQ Shares Outstanding Without Withholding from Total Shares Granted.
df_rsu4 = df_rsu
df_rsu4 = df_rsu.groupby(["Vest_Date"],as_index=False)["Vest_Shares"].sum()
print(df_rsu4)
df_cal6["BOQ Shares Outstanding Without Withholding"]=df_cal6.apply(lambda x: df_rsu4.loc[(df_rsu4.Vest_Date<=x.Qtr_Start),"Vest_Shares"].sum(),axis=1)
df_cal6["Unvested Shares BOQ"]=df_cal6["Total Shares Granted"]-df_cal6["BOQ Shares Outstanding Without Withholding"]
print(df_cal6)

#Creates another dataframe, cumulativebuybackshares, which sums column Buy Back Shares from amorttable by quarter-end dates. Sets df_cal6's Buy Back Shares column to cumulativebuybackshares.
cumulativebuybackshares = amorttable.groupby(["Qtr_End"],as_index=False)["Buy Back Shares"].sum()
print(cumulativebuybackshares)
df_cal6["Buy Back Shares"]=cumulativebuybackshares["Buy Back Shares"]
print(df_cal6)

#Calculates Diluted Shares by subtracting Buy Back Shares from Unvested Shares BOQ.
df_cal6["Diluted Shares"]=df_cal6["Unvested Shares BOQ"]-df_cal6["Buy Back Shares"]
print(df_cal6)

#Creates dataframe finaloutput equivalent to df_cal6's Qtr_End and Diluted Shares columns. Adds Total Weighted Average column equivalent to sharesoutstanding's Total Weighted Average column.
finaloutput = df_cal6[["Qtr_End","Diluted Shares"]]
print(finaloutput)
finaloutput["Total Weighted Average"]=sharesoutstanding["Total Weighted Average"]

#Calculates Total Diluted Shares Outstanding by summing Diluted Shares and Total Weighted Average.
finaloutput["Total Diluted Shares Outstanding"]=finaloutput["Diluted Shares"]+finaloutput["Total Weighted Average"]
print(finaloutput)

