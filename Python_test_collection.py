import openpyxl
import xlrd
import xlsxwriter 

##################################################################################################################

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
    
    def saveCardInfo(self):

    	# create a new workbook-only once
        # if final_magic sheet does not exist create one: 
        new_workbook = xlsxwriter.Workbook('final_magic.xlsx')
        # add one sheet and assign a name-only once
        new_sheet = new_workbook.add_worksheet('NewZen')
        
        i = 0

        # adding header -only once
        for item in cardInfo:
            xwrite = new_sheet.write(1,i,cardInfo[i])
            i += 1
        
        new_workbook.close()
        print('Card saved to new excel file')

    def setHeader(self):
        header_list = ['Card name','Color','Rarity','Foil', 'Number','Special','Location','Price']
        
        open_workbook = openpyxl.load_workbook('final_magic.xlsx')
        open_sheet = open_workbook.get_sheet_by_name('NewZen')

        for item in range(len(header_list)):
             = open_sheet.append([header_list[item]])
            open_workbook.save('final_magic.xlsx') 
        
#             for i in range(row):
#    ws_write.append([datalist[i]])
# wb.save(filename='data.xlsx')
        print('finished setting first row, once for all')


###################################################################################################################
from lxml import html
import requests

class WebScraperMTG:
    
    # Private members
    # self.webpage
    # self.card_list
    # self.price_list
    
    # Private methods
    # setInformationLists()
    
    # Public methods
    # __init__(webpage)
    # getWebInformationList(cardname) - returns a list
    
    def __init__(self,webpage):
  
        self.webpage = webpage
        self.price_list = []
        self.card_list = []
        self.setInformationLists() 
        
    def setInformationLists(self):

        page = requests.get(self.webpage)
        tree = html.fromstring(page.content)
        
        card_list = tree.xpath('//a[@data-full-image]/text()')

        size = int(len(card_list) / 2)
        self.card_list = card_list[0:size]
        print(self.card_list)
        
        price_list = tree.xpath('//td[@class="text-right"]/text()')

        price_list = [x for x in price_list if x != '\n']
        size = int(len(price_list)/2)
        self.price_list = []
        for i in range(0,size,3):
            value = price_list[i]
            value = value[1:-1]
            self.price_list.append(value)
    
#        self.daily_change_list = []
#        for j in range(1,size,3):
#            value = price_list[j]
#            value = value[1:-1]
#            self.daily_change_list.append(value)
#        newSize = int(len(self.daily_change_list))
#        self.daily_change_list = self.daily_change_list[0:newSize]
#        print(newSize)
#        print(self.daily_change_list)      
        
#        self.weekly_change_list = [] for k in range(2,size,3):     value = price_list[j]     value = value[1:-1]
#        self.weekly_change_list.append(value) newSize = int(len(self.weekly_change_list)) self.weekly_change_list =
#        self.weekly_change_list[0:newSize] print(newSize) print(self.weekly_change_list)

    def getWebInformationList(self,cardname):
        # Loop through self.card_list until you find the element that matches cardname
        # Get value of self.price_list at same element number
        
    def getWebPriceList(self):
        return self.price_list

###################################################################################################################

class CardSummary:
    # Creates output file (create the correct header columns)
    # Writes each merged combined line from card collection and web page to output file
    # Saves output file


###################################################################################################################

# Card Collection Test Suite  
oCC = CardCollection('/home/robert/Python/Magic/MTG_Collection_4_20_16.xlsx')
numSheets = oCC.getNumSheets()
numRows = oCC.getNumRows(46)
sheetName = oCC.getSheetname(46)
cardName = oCC.getCardname(46,1)
cardInfo = oCC.getAllCardInfo(46,1)
#output_file = oCC.saveCardInfo()

oWS = WebScraperMTG('http://www.mtggoldfish.com/index/ZEN#paper')
priceList = oWS.getWebPriceList()

print("Price List from object: ", priceList)

print('Total number of SHEETS in this file is:', numSheets)
print('Total number of ROWS in this sheet is:',numRows)
print('The SHEET NAME found is:',sheetName)
print('The CARD NAME is:',cardName)
print('The list of all required information is:', cardInfo)

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
