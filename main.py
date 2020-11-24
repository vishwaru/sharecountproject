import pandas as pd
import numpy as np

#Function returns calendar details on the current quarter and the next 4 quarters (Year, Quarter, Qtr_Ticker, Qtr_Start, Qtr_End)
def adbefiscalcalendar():
    financialcalendar = r"\\sjshare\finance\Treasury\Financial_Database\ADOBE_FINANCIAL_CALENDAR.xlsx"
    sheet = "ADBE_cal"
    df_cal = pd.read_excel(financialcalendar, sheet_name=sheet)
    today = pd.Timestamp.today()
    df_cal = df_cal[df_cal["Per_End"] >= today]
    df_cal = df_cal[["Year", "Quarter", "Qtr_Ticker", "Qtr_Start", "Qtr_End"]]
    df_cal = df_cal.drop_duplicates()
    df_cal = df_cal.reset_index(drop=True)
    df_cal = df_cal.head(5)
    return df_cal

#Function returns details of every grant number at each of its vest dates (Grant_Number, Grant_Date, Plan, Vest_Shares, Grant_Price, Vest_Date, Country, Tax_Rate, SHS (Shares) W/H (Withheld), Vested SHS Post-W/H (Post-Withholding))
def rsutable():
    equitydatabase = r"\\sjshare\finance\Treasury\Financial_Database\EQUITY_COMP_FULL.xlsm"
    sheet2 = "RSU"
    df_rsu = pd.read_excel(equitydatabase, sheet_name = sheet2)
    df_rsu["SHS W/H"] = df_rsu["SHS W/H"].apply(np.ceil)
    df_rsu["Vested SHS Post-W/H"] = df_rsu["Vest_Shares"] - df_rsu["SHS W/H"]
    df_rsu = df_rsu.drop(columns=["Column1"])
    return df_rsu

#Function returns the Final Vest Date, Grant Date, Shares Granted, Grant Price, Total Expense Days, and Total Expense for each unique grant number (row)
def grantinformation():
    grantinfo = rsutable().groupby(["Grant_Number"]).agg(Vest_Date = ("Vest_Date",'max'), Grant_Date = ("Grant_Date", 'max'), Vest_Shares = ("Vest_Shares", 'sum'), Grant_Price = ("Grant_Price", 'max'))
    grantinfo["Total Expense Days"] = ((grantinfo["Vest_Date"] - grantinfo["Grant_Date"]).dt.days) + 1
    grantinfo = grantinfo.rename(columns={"Vest_Shares": "Shares Granted", "Vest_Date":"Final Vest_Date"})
    grantinfo["Total Expense"] = grantinfo["Grant_Price"] * grantinfo["Shares Granted"]
    return grantinfo

#Function returns the Amortized Expense, Average Unamortized Expense, BOQ Unamortized Expense, Buy Back Shares, EOQ Unamortized Expense, Est Avg Share Price, and Expense_Days for each of the five quarters for each unique grant number (row)
def amortizationtable():
    d = {}
    df_calstr = adbefiscalcalendar().astype(str)
    for i in range(0,len(adbefiscalcalendar())):
        d[df_calstr.iloc[i,0] + "Q" + df_calstr.iloc[i,1]] = pd.DataFrame(grantinformation())
        amorttable = pd.concat(d, axis=1)
    for i in range(0,len(adbefiscalcalendar())):
        amorttable[df_calstr.iloc[i,0] + "Q" + df_calstr.iloc[i,1], "Qtr_End"] = adbefiscalcalendar().iloc[i, 4]
        amorttable[df_calstr.iloc[i, 0] + "Q" + df_calstr.iloc[i, 1], "Expense_Days"] = ((amorttable[df_calstr.iloc[i,0] + "Q" + df_calstr.iloc[i,1], "Qtr_End"]-amorttable[df_calstr.iloc[i,0] + "Q" + df_calstr.iloc[i,1], "Grant_Date"]).dt.days)+1
        amorttable[df_calstr.iloc[i, 0] + "Q" + df_calstr.iloc[i, 1], "Expense_Days"] = amorttable[[(df_calstr.iloc[i, 0] + "Q" + df_calstr.iloc[i, 1], "Expense_Days"),(df_calstr.iloc[i, 0] + "Q" + df_calstr.iloc[i, 1], "Total Expense Days")]].min(axis=1)
        amorttable[df_calstr.iloc[i, 0] + "Q" + df_calstr.iloc[i, 1], "Amortized Expense"] = amorttable[df_calstr.iloc[i, 0] + "Q" + df_calstr.iloc[i, 1], "Total Expense"] * (amorttable[df_calstr.iloc[i, 0] + "Q" + df_calstr.iloc[i, 1], "Expense_Days"]/amorttable[df_calstr.iloc[i, 0] + "Q" + df_calstr.iloc[i, 1], "Total Expense Days"])
        amorttable[df_calstr.iloc[i, 0] + "Q" + df_calstr.iloc[i, 1], "Qtr_Start"] = adbefiscalcalendar().iloc[i, 3]
        amorttable = amorttable.sort_index(axis=1)
    subamorttable = amorttable[df_calstr.iloc[0, 0] + "Q" + df_calstr.iloc[0, 1]]
    subamorttable2 = subamorttable[subamorttable["Grant_Date"]>= subamorttable["Qtr_Start"]]
    subamorttable2["BOQ Unamortized Expense"] = subamorttable2["Total Expense"]
    subamorttable2["EOQ Unamortized Expense"] = subamorttable2["Total Expense"] - subamorttable2["Amortized Expense"]
    subamorttable2["Average Unamortized Expense"] = ((subamorttable2["BOQ Unamortized Expense"] + subamorttable2["EOQ Unamortized Expense"]) / (2)) * ((subamorttable2["Expense_Days"]) / (((subamorttable2["Qtr_End"] - subamorttable2["Qtr_Start"]).dt.days) + 1))
    subamorttable3 = subamorttable[subamorttable["Grant_Date"] < subamorttable["Qtr_Start"]]
    subamorttable3["Prev Qtr Expense Days"] = (((subamorttable3["Qtr_Start"]-pd.Timedelta(1,unit='d'))-(subamorttable3["Grant_Date"])).dt.days)+1
    subamorttable3["Prev Qtr Expense Days"] = subamorttable3[["Total Expense Days", "Prev Qtr Expense Days"]].min(axis=1)
    subamorttable3["BOQ Unamortized Expense"] = (subamorttable3["Total Expense"]) - (subamorttable3["Total Expense"] * (subamorttable3["Prev Qtr Expense Days"] / subamorttable3["Total Expense Days"]))
    subamorttable3["EOQ Unamortized Expense"] = subamorttable3["Total Expense"] - subamorttable3["Amortized Expense"]
    subamorttable3["Average Unamortized Expense"] = ((subamorttable3["BOQ Unamortized Expense"] + subamorttable3["EOQ Unamortized Expense"]) / (2))
    subamorttable2and3 = pd.concat([subamorttable2.drop(columns=["Final Vest_Date", "Grant_Date", "Grant_Price", "Qtr_End", "Qtr_Start", "Shares Granted", "Total Expense", "Total Expense Days"]), subamorttable3.drop(columns=["Final Vest_Date", "Grant_Date", "Grant_Price", "Qtr_End", "Qtr_Start", "Shares Granted", "Total Expense", "Total Expense Days", "Prev Qtr Expense Days"])], axis=0, ignore_index=False)
    subamorttable2and3.columns = pd.MultiIndex.from_product([[df_calstr.iloc[0, 0] + "Q" + df_calstr.iloc[0, 1]], subamorttable2and3.columns])
    subamorttable4 = amorttable.drop(columns=[df_calstr.iloc[0, 0] + "Q" + df_calstr.iloc[0, 1]])
    for i in range(1,len(adbefiscalcalendar())):
        subamorttable4[df_calstr.iloc[i, 0] + "Q" + df_calstr.iloc[i, 1], "Prev Qtr Expense Days"] = (((subamorttable4[df_calstr.iloc[i, 0] + "Q" + df_calstr.iloc[i, 1], "Qtr_Start"]-pd.Timedelta(1,unit='d'))-(subamorttable4[df_calstr.iloc[i, 0] + "Q" + df_calstr.iloc[i, 1], "Grant_Date"])).dt.days)+1
        subamorttable4[df_calstr.iloc[i, 0] + "Q" + df_calstr.iloc[i, 1], "Prev Qtr Expense Days"] = subamorttable4[[(df_calstr.iloc[i, 0] + "Q" + df_calstr.iloc[i, 1], "Prev Qtr Expense Days"),(df_calstr.iloc[i, 0] + "Q" + df_calstr.iloc[i, 1], "Total Expense Days")]].min(axis=1)
        subamorttable4[df_calstr.iloc[i, 0] + "Q" + df_calstr.iloc[i, 1], "BOQ Unamortized Expense"] = (subamorttable4[df_calstr.iloc[i, 0] + "Q" + df_calstr.iloc[i, 1], "Total Expense"]) - (subamorttable4[df_calstr.iloc[i, 0] + "Q" + df_calstr.iloc[i, 1], "Total Expense"] * (subamorttable4[df_calstr.iloc[i, 0] + "Q" + df_calstr.iloc[i, 1], "Prev Qtr Expense Days"]/subamorttable4[df_calstr.iloc[i, 0] + "Q" + df_calstr.iloc[i, 1], "Total Expense Days"]))
        subamorttable4[df_calstr.iloc[i, 0] + "Q" + df_calstr.iloc[i, 1], "EOQ Unamortized Expense"] = subamorttable4[df_calstr.iloc[i, 0] + "Q" + df_calstr.iloc[i, 1], "Total Expense"] - subamorttable4[df_calstr.iloc[i, 0] + "Q" + df_calstr.iloc[i, 1], "Amortized Expense"]
        subamorttable4[df_calstr.iloc[i, 0] + "Q" + df_calstr.iloc[i, 1], "Average Unamortized Expense"] = ((subamorttable4[df_calstr.iloc[i, 0] + "Q" + df_calstr.iloc[i, 1], "BOQ Unamortized Expense"] + subamorttable4[df_calstr.iloc[i, 0] + "Q" + df_calstr.iloc[i, 1], "EOQ Unamortized Expense"])/(2))
        subamorttable4 = subamorttable4.drop(columns=[(df_calstr.iloc[i, 0] + "Q" + df_calstr.iloc[i, 1], "Final Vest_Date"), (df_calstr.iloc[i, 0] + "Q" + df_calstr.iloc[i, 1], "Grant_Date"), (df_calstr.iloc[i, 0] + "Q" + df_calstr.iloc[i, 1], "Grant_Price"), (df_calstr.iloc[i, 0] + "Q" + df_calstr.iloc[i, 1], "Prev Qtr Expense Days"), (df_calstr.iloc[i, 0] + "Q" + df_calstr.iloc[i, 1], "Qtr_End"), (df_calstr.iloc[i, 0] + "Q" + df_calstr.iloc[i, 1], "Qtr_Start"), (df_calstr.iloc[i, 0] + "Q" + df_calstr.iloc[i, 1], "Shares Granted"), (df_calstr.iloc[i, 0] + "Q" + df_calstr.iloc[i, 1], "Total Expense"), (df_calstr.iloc[i, 0] + "Q" + df_calstr.iloc[i, 1], "Total Expense Days")])
        subamorttable4 = subamorttable4.sort_index(axis=1)
    amorttable = pd.concat([subamorttable2and3, subamorttable4], axis=1).sort_index(axis=1)
    currentqavgshareprice = r"\\sjshare\finance\Treasury\Financial_Database\ADBE Historicals.xlsm"
    sheetininterest = "CurrentQAvgSharePrice"
    cqavgshareprice = pd.read_excel(currentqavgshareprice, sheet_name=sheetininterest, usecols = "C:C")
    for i in range(1, len(adbefiscalcalendar())):
        cqavgshareprice.loc[i,"Current Quarter Est Avg Share Price"]=cqavgshareprice.loc[i-1,"Current Quarter Est Avg Share Price"]*((1.1)**(0.25))
    for i in range(0, len(adbefiscalcalendar())):
        amorttable[df_calstr.iloc[i, 0] + "Q" + df_calstr.iloc[i, 1], "Est Avg Share Price"] = cqavgshareprice.loc[i,"Current Quarter Est Avg Share Price"]
        amorttable[df_calstr.iloc[i, 0] + "Q" + df_calstr.iloc[i, 1], "Buy Back Shares"] = (amorttable[df_calstr.iloc[i, 0] + "Q" + df_calstr.iloc[i, 1], "Average Unamortized Expense"])/(amorttable[df_calstr.iloc[i, 0] + "Q" + df_calstr.iloc[i, 1], "Est Avg Share Price"])
        amorttable = amorttable.sort_index(axis=1)
    return amorttable

#Function returns the number of Shares Outstanding, BOQ Shares Outstanding, Qtr_Weight, Total Weighted Average, and Unvested Shares BOQ for each of the 5 quarters
def sharestable():
    df_rsu2 = rsutable().groupby(["Vest_Date"], as_index=False).agg(Vested_SHS_Post_WH = ("Vested SHS Post-W/H", 'sum'), Vest_Shares = ("Vest_Shares", 'sum'))
    df_rsu2 = df_rsu2[df_rsu2["Vest_Date"]>= adbefiscalcalendar().iloc[0,3]]
    shstable = adbefiscalcalendar()[["Qtr_Start","Qtr_End"]]
    shstable["Shares Outstanding"] = shstable.apply(lambda x: df_rsu2.loc[(df_rsu2.Vest_Date <= x.Qtr_End), "Vested_SHS_Post_WH"].sum(), axis=1)
    shstable["BOQ Shares Outstanding"] = shstable.apply(lambda x: df_rsu2.loc[(df_rsu2.Vest_Date <= x.Qtr_Start), "Vested_SHS_Post_WH"].sum(), axis=1)
    qtr_end = []
    for i in df_rsu2["Vest_Date"]:
        qtr_end.append((shstable.loc[(shstable["Qtr_Start"]<=i)&(shstable["Qtr_End"]>=i),"Qtr_End"]))
    df_rsu2["Qtr_End"] = pd.concat(qtr_end,ignore_index=True)
    df_rsu2 = pd.merge(df_rsu2, adbefiscalcalendar()[["Qtr_Start","Qtr_End"]], on="Qtr_End", how="inner")
    df_rsu2["Qtr_Weight"] = df_rsu2["Vested_SHS_Post_WH"] * ((((df_rsu2["Qtr_End"] - df_rsu2["Vest_Date"]).dt.days)+1)/(((df_rsu2["Qtr_End"]-df_rsu2["Qtr_Start"]).dt.days)+1))
    shstable = pd.merge(shstable, (df_rsu2.groupby(["Qtr_End"], as_index=False)["Qtr_Weight"].sum()), on="Qtr_End", how="left").fillna(0)
    shstable["Total Weighted Average"] = shstable["BOQ Shares Outstanding"] + shstable["Qtr_Weight"]
    shstable["Unvested Shares BOQ"] = (grantinformation()["Shares Granted"].sum()) - (shstable.apply(lambda x: df_rsu2.loc[(df_rsu2.Vest_Date<=x.Qtr_Start),"Vest_Shares"].sum(),axis=1))
    pd.options.display.float_format = "{:.6f}".format
    shstable = shstable.drop(columns=["Qtr_Start","Qtr_End"]).transpose()
    list = []
    for i in range(0,len(adbefiscalcalendar())):
        list.append((adbefiscalcalendar().astype(str)).iloc[i, 0] + "Q" + (adbefiscalcalendar().astype(str)).iloc[i, 1])
    shstable.columns = list
    return shstable

#Function returns the Total Weighted Average, Diluted Shares, and Total Diluted Shares Outstanding for each of the 5 quarters
def finaloutput():
    finaloutputtable = sharestable().transpose()
    for i in range(0,len(adbefiscalcalendar())):
        finaloutputtable.loc[((adbefiscalcalendar().astype(str)).iloc[i, 0]) + "Q" + ((adbefiscalcalendar().astype(str)).iloc[i, 1]), "Buy Back Shares"] = amortizationtable()[((adbefiscalcalendar().astype(str)).iloc[i, 0]) + "Q" + ((adbefiscalcalendar().astype(str)).iloc[i,1]), "Buy Back Shares"].sum()
        finaloutputtable.loc[((adbefiscalcalendar().astype(str)).iloc[i, 0]) + "Q" + ((adbefiscalcalendar().astype(str)).iloc[i, 1]), "Diluted Shares"] = finaloutputtable.loc[((adbefiscalcalendar().astype(str)).iloc[i, 0]) + "Q" + ((adbefiscalcalendar().astype(str)).iloc[i, 1]), "Unvested Shares BOQ"] - finaloutputtable.loc[((adbefiscalcalendar().astype(str)).iloc[i, 0]) + "Q" + ((adbefiscalcalendar().astype(str)).iloc[i, 1]), "Buy Back Shares"]
        finaloutputtable.loc[((adbefiscalcalendar().astype(str)).iloc[i, 0]) + "Q" + ((adbefiscalcalendar().astype(str)).iloc[i, 1]), "Total Diluted Shares Outstanding"] = finaloutputtable.loc[((adbefiscalcalendar().astype(str)).iloc[i, 0]) + "Q" + ((adbefiscalcalendar().astype(str)).iloc[i, 1]), "Diluted Shares"] + finaloutputtable.loc[((adbefiscalcalendar().astype(str)).iloc[i, 0]) + "Q" + ((adbefiscalcalendar().astype(str)).iloc[i, 1]), "Total Weighted Average"]
    finaloutputtable = finaloutputtable.drop(columns=["Shares Outstanding", "BOQ Shares Outstanding", "Qtr_Weight", "Unvested Shares BOQ", "Buy Back Shares"]).transpose()
    return finaloutputtable

#Calls function finaloutput() under var name finaloutputtable and prints the results
finaloutputtable = finaloutput()
print(finaloutputtable)
