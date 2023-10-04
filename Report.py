#pip install pandas
#pip install openpyxl
#pip install XlsxWriter

import pandas as pd

tableResults = pd.read_excel("Results.xlsx")
tableSiteList = pd.read_excel("SiteList.xlsx")

tableReport = pd.merge(tableResults, tableSiteList, on=["Site Name","Site ID", "Equipment", "Year"]) #Find the intersection between tables
tableReport = tableReport[["Site Name", "Site ID", "State", "Equipment", "Signal (%)", "Quality (0-10)", "Mbps"]] #rearranges the report table only with what is requested and in order
tableReport = tableReport.sort_values(by="State", ascending=True) #Sorts the rows in alphabetical order from the "States" column


#Check the alerted sites
results_Alert = tableResults[tableResults["Alerts"] == "Yes"]
sitesResultAlert = results_Alert["Site Name"]
print("Sites with active alerts: ")
print(sitesResultAlert.to_string(index=False)) #Print as a string (without pandas index)

#Check the sites with 0 quality
results_zeroQuality = tableResults[tableResults["Quality (0-10)"] == 0]
sitesResultZeroQuality = results_zeroQuality["Site Name"]
print("\nSites with zero quality: ")
print(sitesResultZeroQuality.to_string(index=False)) #Print as a string (without pandas index)

#Check the sites with more then 80 Mbps
results_moreThen80Mbps = tableResults[tableResults["Mbps"] > 80]
sitesResultMoreThen80Mbps = results_moreThen80Mbps["Site Name"]
print("\nSites with more then 80 Mbps: ")
print(sitesResultMoreThen80Mbps.to_string(index=False)) #Print as a string (without pandas index)

#Check the sites with more then 80 Mbps
results_lessThen10Mbps = tableResults[tableResults["Mbps"] < 10]
sitesResultLessThen10Mbps = results_lessThen10Mbps["Site Name"]
print("\nSites with less then 10 Mbps: ")
print(sitesResultLessThen10Mbps.to_string(index=False)) #Print as a string (without pandas index)


#Writing an Excel file with many tabs

fileName = "Table_Report.xlsx"

with pd.ExcelWriter(fileName, engine='xlsxwriter') as writer:
    tableReport.to_excel(writer, sheet_name="Report Site List", index = False) # Report from Site List
    sitesResultAlert.to_excel(writer, sheet_name="Alert", index = False)
    sitesResultZeroQuality.to_excel(writer, sheet_name="Zero Quality", index = False)
    sitesResultMoreThen80Mbps.to_excel(writer, sheet_name="Over 80 Mbps", index = False)
    sitesResultLessThen10Mbps.to_excel(writer, sheet_name="Below 10 Mbps", index = False)
