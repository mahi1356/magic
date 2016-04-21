import openpyxl
import os

class CardCollection:
   
    # Class variables
    # self.collectionFile - This is the name for the excel spreadsheet
    
    def __init__(self,collectionFile):
        #os.chdir('/home/robert/Python/Magic/')
        os.chdir('/Users/mity/mypy')
        
        self.collectionFile = collectionFile
        self.sourceWorkbook = openpyxl.load_workbook(collectionFile)
        
    def getNumSheets(self):
        return 67
    
    def getNumRows(self):
        return 324

    def getSheetname(self,sheetIndex):
        return 'ZEN'
    
    def getCardname(self,sheetIndex,rowIndex):
        sourceSheetName = 'ZEN'
        sourceSheet = self.sourceWorkbook.get_sheet_by_name(sourceSheetName)
        cardName = sourceSheet.cell(row = rowIndex+1,column=2).value

        return cardName
    
    def getAllCardInfo(self,sheetIndex,rowIndex):
        # I want to use dictionary
        cardInfo = {}
        cardInfo.append{'Card' :'Bala Ged Thief'}
        # read about dictionary then add items to dictionary 
        cardInfo.append(['Color' : 'B'])
        cardInfo.append(['Rarity' :' Rare'])
        cardInfo.append(['Foil' : 'No'])
        return cardInfo
    
    def getCardPrice(self):
        return '2.00'
    
    def saveCardInfo(self):
        return TRUE

# Card Collection Test Suite  

oCC = CardCollection('MTG_Collection_4_20_16.xlsx')

numSheets = oCC.getNumSheets()
numRows = oCC.getNumRows()
sheetName = oCC.getSheetname(47)
cardName = oCC.getCardname(47,1)
cardInfo = oCC.getAllCardInfo(47,1)


print(numSheets)
print(numRows)
print(sheetName)
print(cardName)
print(cardInfo)

if numSheets == 67:
    print("numSheets Check = TRUE")
else:
    print("numSheets Check = FALSE")
    
if numRows == 324:
    print("numRows Check = TRUE")
else:
    print("numRows Check = FALSE")

if sheetName == "ZEN":
    print("sheetName Check = TRUE")
else:
    print("sheetName Check = FALSE")

if cardName == "Bala Ged Thief":
    print("cardName Check = TRUE")
else:
    print("cardName Check = FALSE")
# Mitra - Write test for cardInfo - should return columns Card, Color, Rarity, Foil, 
#Special, Number and Location from   spreadsheet

cardName = oCC.getCardname(47,5)
cardInfo = oCC.getAllCardInfo(47,5)

print(cardName)
if cardName == "Bloodchief Ascension":
    print("Foil cardname Check = TRUE")
else:
    print("Foil cardname Check = FALSE")
