# Main program

oCC = CardCollection('/home/robert/Python/Magic/MTG_Collection_4_20_16.xlsx')
oCS = CardSummary('final_magic.xlsx')
numSheets = oCC.getNumSheets()
# this is a correct syntax but the current functionality is not 
# advance enough to use this loop--->for sheetIndex in range (numSheets-1):
sheetName = oCC.getSheetname(sheetIndex)
# Make webpage url using sheetName

sheetIndex = 1   # change this after sheet loop is writen 

oWS = WebScraperMTG('http://www.mtggoldfish.com/index/ZEN#paper')
numRows = oCC.getNumRows(sheetIndex)
# Loop through all rows on current sheet
for rowIndex in range(numRows-1):
    cardName = oCC.getCardname(sheetIndex,rowIndex)
    cardInfo = oCC.getAllCardInfo(sheetIndex,rowIndex)
    webpageInfo = oWS.getWebInformationList(cardName)
    summaryList = cardInfo + webpageInfo
    summaryList.append(sheetName)
    oCS.writeSummaryRow(summaryList)
# is it better to have savefile in main or inside card summary- why?

oSummary.saveFile()
