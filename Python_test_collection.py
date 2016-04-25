# Main program
import openpyxl
import xlrd
import os

os.chdir('/Users/mity/mypy')
#os.chdir('/home/robert/Python/Magic/')

# Open Collection by creating a CardCollection object (oCC)
# Find number of sheets

# Loop through all sheets (using number of sheets)
#for sheetIndex=0 to numSheets
    # Get name of the current sheet
    # Open web page to scrape the data from (by creating a MTG_Webpage object)
    # Get number of rows of current sheet
    # Loop through all rows on current sheet
        # Get cardInfo line from oCC
        # Get card name for current row
        # Go to web site and get the price
        # price = oWebPage.getCardPrice(cardName)
        # Append price to cardInfo line to create summaryLine for outputfile
        # Send summaryLine to the OutputFile object
# Save final summary fileimport openpyxl

class CardCollection:

    def __init__(self,collectionFile):
  
    # Mitra to create member variable self.workingDir to hold directory to use
    # Could either ask who's running the program
    # Mitra also to create member variable self.collection to hold filename and set it here
        self.collectionFile = collectionFile
        self.sourceWorkbook = openpyxl.load_workbook(collectionFile)

    def getNumSheets(self):
       
        book = xlrd.open_workbook('MTG_Collection_4_20_16.xlsx') 
        # delete when member variable self.collection created
        numb_sheets=book.nsheets
        return numb_sheets

    def getNumRows(self, sheetIndex):
        return 324

    def getSheetname(self,sheetIndex):
        return 'ZEN'
    
    def getCardname(self,sheetIndex,rowIndex):
        sourceSheetName = 'ZEN'
        sourceSheet = self.sourceWorkbook.get_sheet_by_name(sourceSheetName)
        cardName = sourceSheet.cell(row = rowIndex+1,column=2).value
        return cardName
    
    def getAllCardInfo(self,sheetIndex,rowIndex):
    
        cardInfo = []
        cardInfo.append('Bala Ged Thief')
        cardInfo.append('B')
        cardInfo.append('Rare')
        cardInfo.append('N')
        cardInfo.append('N')
        cardInfo.append('1')
        cardInfo.append('Zendikar - Storage Box')
      
        return cardInfo
    
    def getCardPrice(self):
        return '2.00'
    
    def saveCardInfo(self):
        return TRUE

# Card Collection Test Suite  

oCC = CardCollection('MTG_Collection_4_20_16.xlsx')

numSheets = oCC.getNumSheets()
numRows = oCC.getNumRows(47)
sheetName = oCC.getSheetname(47)
cardName = oCC.getCardname(47,1)
cardInfo = oCC.getAllCardInfo(47,1)


print('Total number of SHEETS in this file is:', numSheets)
print('Total number of ROWS in this sheet is:',numRows)
print('The SHEET NAME found is:',sheetName)
print('The CARD NAME is:',cardName)
print('The list of all required information is:', cardInfo)


if numSheets == 69:
    print("numSheets Check = TRUE")
else:
    print("numSheets Check = FALSE")
    ###print("It should be"+ numSheets.numb_sheets)
    
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

if cardInfo[0] == 'Bala Ged Thief':
    print("Cardname Check from list = TRUE")
else:
    print("Cardname Check from list = FALSE")

if cardInfo[1] == 'B':
     print("Color Check from list = TRUE")
else:
     print("Color Check from list = FALSE") 

if cardInfo[2] == 'Rare':
    print("Rarity Check from list= TRUE")
else:
    print("Rarity Check from list=  FALSE")

if cardInfo[3] == 'N':   
    print("Foil Check from list = TRUE")
else:
    print("Foil Check from list = FALSE")

if cardInfo[4] == 'N':   
    print("Special Check from list = TRUE")
else:
    print("Special Check from list = FALSE")

if cardInfo[5] == '1':  
    print("Number Check from list = TRUE")
else:
    print("Number from list = FALSE")

if cardInfo[6] == 'Zendikar - Storage Box':   
    print("Location Check from list = TRUE")
else:
    print("Location Check from list = FALSE")
cardName = oCC.getCardname(47,5)
cardInfo = oCC.getAllCardInfo(47,5)

print(cardName)
if cardName == "Bloodchief Ascension":
    print("cardname Check = TRUE")
else:
    print("cardname Check = FALSE")
