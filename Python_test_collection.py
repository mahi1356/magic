# Main program
import openpyxl
import xlrd
import xlsxwriter 
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
        self.sourceWorkbook = xlrd.open_workbook(self.collectionFile) 

    def getNumSheets(self):
        # delete when member variable self.collection created
        numb_sheets = self.sourceWorkbook.nsheets
        return numb_sheets

    def getNumRows(self, sheetIndex):
        sheet_dest = self.sourceWorkbook.sheet_by_index(sheetIndex)
        numb_rows = sheet_dest.nrows - 1
        return numb_rows

    def getSheetname(self,sheetIndex):
        sh_names_list= self.sourceWorkbook.sheet_names()
        sh_name = sh_names_list[sheetIndex]
        return sh_name
    
    def getCardname(self,sheetIndex,rowIndex):
     
        # below will get a list of sheets  using sheet_by_name  
        #sh_names= self.sourceWorkbook.sheet_names()
        #my_sheet = self.sourceWorkbook.sheet_by_name(sh_names[rowIndex])
       
        #below is shorter than above and more clear
        my_sheet = self.sourceWorkbook.sheet_by_index(sheetIndex)
        my_cell = my_sheet.cell(rowIndex,1).value
        #Mitra noticed that return variable could be different than what is name in main program
        return my_cell
    
    def getAllCardInfo(self,sheetIndex,rowIndex):
         # mitra to finish this section 
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

    	# create a new workbook-only once
        # if final_magic sheet does not exist create one: 
        new_workbook = xlsxwriter.Workbook('final_magic.xlsx')
        # add one sheet and assign a name-only once
        new_sheet = new_workbook.add_worksheet('NewZen')
        
        # for j in range(0,8):
        #     header_sheet= new_sheet.write (0,j, 'Card name')
        #     j += 1
        i = 0

        # adding header -only once
        for item in cardInfo:
            xwrite = new_sheet.write(1,i,cardInfo[i])
            i += 1
        
        new_workbook.close()
   

# Card Collection Test Suite  
oCC = CardCollection('MTG_Collection_4_20_16.xlsx')
numSheets = oCC.getNumSheets()
numRows = oCC.getNumRows(46)
sheetName = oCC.getSheetname(46)
cardName = oCC.getCardname(46,1)
cardInfo = oCC.getAllCardInfo(46,1)
output_file = oCC.saveCardInfo()

print('Total number of SHEETS in this file is:', numSheets)
print('Total number of ROWS in this sheet is:',numRows)
print('The SHEET NAME found is:',sheetName)
print('The CARD NAME is:',cardName)
print('The list of all required information is:', cardInfo)
print('Card saved to new excel file')

# if numSheets == 67:
#     print("numSheets Check = TRUE")
# else:
#     ##print("numSheets Check = FALSE")
#     print("numsheet reported is FALSE, expected value is" , numSheets)
    
# if numRows == 323:
#     print("numRows Check = TRUE")
# else:
#     print("numRows Check = FALSE")

# if sheetName == "ZEN":
#     print("sheetName Check = TRUE")
# else:
#     print("sheetName Check = FALSE")

# if cardName == "Bala Ged Thief":
#     print("cardName Check = TRUE")
# else:
#     print("cardName Check = FALSE")

# if cardInfo[0] == 'Bala Ged Thief':
#     print("Cardname Check from list = TRUE")
# else:
#     print("Cardname Check from list = FALSE")

# if cardInfo[1] == 'B':
#      print("Color Check from list = TRUE")
# else:
#      print("Color Check from list = FALSE") 

# if cardInfo[2] == 'Rare':
#     print("Rarity Check from list= TRUE")
# else:
#     print("Rarity Check from list=  FALSE")

# if cardInfo[3] == 'N':   
#     print("Foil Check from list = TRUE")
# else:
#     print("Foil Check from list = FALSE")

# if cardInfo[4] == 'N':   
#     print("Special Check from list = TRUE")
# else:
#     print("Special Check from list = FALSE")

# if cardInfo[5] == '1':  
#     print("Number Check from list = TRUE")
# else:
#     print("Number from list = FALSE")

# if cardInfo[6] == 'Zendikar - Storage Box':   
#     print("Location Check from list = TRUE")
# else:
#     print("Location Check from list = FALSE")
# cardName = oCC.getCardname(47,5)
# cardInfo = oCC.getAllCardInfo(47,5)

# print(cardName)
# if cardName == "Bloodchief Ascension":
#     print("cardname Check = TRUE")
# else:
#     print("cardname Check = FALSE")
