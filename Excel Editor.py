import pandas as pd
import os
import time
import sys

class moneyKeep():
    # Database Selection
    bindedType = '.xlsx'
    selectedDB = None
    dbDir = None
    selectedDBName = None
    loop1 = 0
    def __init__(self,dbdir):
        self.dbDir = dbdir
        self.dbSelect()

    def clearConsole(self):
        os.system('cls')
        time.sleep(0.5)

    def dbSelect(self):
        loopBoolean1 = True
        dbdict = dict()
        while loopBoolean1:
            dbnum = 1
            try:
                if self.loop1 is 0:
                    print('=' * 10 + 'Select Database you want to handle' + '=' * 10)
                    print('0 . Create New DB')
                    dbli = os.listdir(self.dbDir)
                    dbli = [file for file in dbli if file.endswith(".xlsx")]
                    for dbname in dbli:
                        dbdict[dbnum] = dbname
                        print(dbnum, ".", dbname)
                        dbnum += 1
                    print('=' * 54)
                    dbselect = int(input(">>"))
                    if dbselect == 0:
                        self.loop1 += 1
                        loopBoolean1 = False
                        self.clearConsole()
                        self.createNewDataBaseDirect()
                        self.clearConsole()
                        self.menu()
                    elif dbselect not in dbdict.keys():
                        print("Please type right number")
                    else:
                        loopBoolean1 = False
                else:
                    print('=' * 10 + 'Select Database you want to handle' + '=' * 10)
                    dbli = os.listdir(self.dbDir)
                    dbli = [file for file in dbli if file.endswith(".xlsx")]
                    for dbname in dbli:
                        dbdict[dbnum] = dbname
                        print(dbnum, ".", dbname)
                        dbnum += 1
                    print('=' * 54)
                    dbselect = int(input(">>"))
                    if dbselect not in dbdict.keys():
                        print("Please type right number")
                    else:
                        self.selectedDBName = dbdict[dbselect].split('.')[0]
                        loopBoolean1 = False
            except ValueError:
                print("Warning : Please type value as Inteager")
                time.sleep(2)
                os.system('clear')

        self.clearConsole()
        # 선택한 DB
        self.loop1+=1
        self.selectedDB = pd.read_excel(dbDir + dbdict[dbselect], index_col=0)
        self.selectedDBName = dbdict[dbselect].split('.')[0]
        self.menu()

    def menu(self):
        loopBoolean2 = True
        # menu loop here
        while loopBoolean2:
            menuNum = self.menuSelect()
            if menuNum == 0:
                self.exitProgram()
            elif menuNum == 1:
                self.createNewDataBase()
            elif menuNum == 2:
                self.viewSelectDB()
            elif menuNum == 3:
                select = input("Do you want to save changed data?[y/n] : ")
                if select == 'Y' or select =='y':
                    self.saveData()
                else:
                    print("Data will not be saved.")
                    self.clearConsole()
                self.dbSelect()
            elif menuNum == 4:
                self.addData()
            elif menuNum == 5:
                self.viewAsPart()
            elif menuNum == 6:
                self.deleteRow()
            elif menuNum == 7:
                self.deleteCol()
            elif menuNum == 8:
                self.saveData()
            else:
                print("You entered Wrong option number")
                self.clearConsole()


    def menuSelect(self):
        print("=" * 8 + " Menu " + "=" * 8)
        print("0 . Exit Program.")
        print("1 . Create New DataBase.")
        print("2 . View entire selected DB.")
        print("3 . Select another DB.")
        print("4 . Add Data.")
        print("5 . View part of selected DB.")
        print("6 . Delete Row.")
        print("7 . Delete Column.")
        print("8 . Save Data.")
        print("=" * 22)
        menuNum = int(input(">>"))
        self.clearConsole()
        return menuNum

    def exitProgram(self):
        print("Program Close...")
        time.sleep(5)
        sys.exit()

    def createNewDataBase(self):
        dbName = input("Enter DB Name : ")
        totalRowCount = int(input("Enter total number of column : "))
        colList = []
        for w in range(0, totalRowCount):
            print("Enter name of column No.", w + 1)
            colName = input('>>')
            colList.append(colName)
        newFrameData = pd.DataFrame(columns=colList)
        newFrameData.to_excel(self.dbDir + dbName + '.xlsx')
        self.clearConsole()
        changeDBYN = input("Do you want change selected db to " + dbName + "?[y / Y or n / N] : ")
        if changeDBYN == 'y' or changeDBYN == 'Y':
            print("Selected database will be change in second...")
            self.selectedDB = pd.read_excel(dbDir + dbName + '.xlsx', index_col=0)
            self.clearConsole()
        elif changeDBYN == 'N' or changeDBYN =='n':
            print("DataBase " + dbName + " generated in dbSaver Directory")
            self.clearConsole()

    def createNewDataBaseDirect(self):
        dbName = input("Enter DB Name : ")
        self.selectedDBName = dbName
        totalRowCount = int(input("Enter total number of column : "))
        colList = []
        for w in range(0, totalRowCount):
            print("Enter name of column No.", w + 1)
            colName = input('>>')
            colList.append(colName)
        newFrameData = pd.DataFrame(columns=colList)
        newFrameData.to_excel(self.dbDir + dbName + '.xlsx')
        self.selectedDB = pd.read_excel(dbDir + dbName + '.xlsx', index_col=0)

    def viewSelectDB(self):
        print(self.selectedDB)

    def selectAntherDB(self):
        self.dbSelect()

    def addData(self):
        loopBoolean = True
        print("=" * 10," Add Data " ,"=" * 10)
        print("Total List of Columns : ",",".join(map(str,list(self.selectedDB.columns))))
        print("=" * 30)
        print("Enter number of Rows you want to add")
        rowCount = int(input(">>"))
        indexList = list(self.selectedDB.index)
        nextIndexNumber = 1
        if len(indexList) is 0:
            pass
        else:
            nextIndexNumber = indexList[-1] + 1
        for q in range(0,rowCount):
            dataSocket = dict()
            for w in list(self.selectedDB.columns):
                print("Enter value of " + w)
                val = input(">>")
                dataSocket[w] = val
            df = pd.DataFrame(
                data=dataSocket,
                index = [nextIndexNumber]
            )
            self.selectedDB = pd.concat([self.selectedDB,df])
            nextIndexNumber += 1
            self.clearConsole()

    def viewAsPart(self):
        print("Enter number of rows you want to see as part.")
        partNum = int(input(">>"))
        self.clearConsole()
        li = []
        for q in range(0, partNum):
            print("Enter index Number")
            indNum = int(input(">>"))
            if indNum - 1 not in range(0,len(self.selectedDB)):
                print("Index number",indNum,"not exist. Ignore this value.")
                self.clearConsole()
                pass
            else:
                li.append(indNum-1)
                self.clearConsole()

        print("=" * 20)
        for w in li:
            print(self.selectedDB.iloc[w])
            print("=" * 20)


    def deleteRow(self):
        print("WARNING : You can't recover your data after you delete Row.")
        select = input("Are you sure you choose this option?[y/n] : ")
        if select == 'Y' or select == 'y':
            self.clearConsole()
            print("Enter Row number you want to delete. There are total ",len(self.selectedDB),"rows.")
            selectRow = int(input(">>"))
            if selectRow -1 not in range(0,len(self.selectedDB)):
                print("There are no index number : ",selectRow)
                self.clearConsole()
                return
            else:
                self.selectedDB = self.selectedDB.drop(selectRow)
                print("Reindexing...")
                reind = []
                for er in range(1,len(self.selectedDB) + 1):
                    reind.append(er)
                self.selectedDB.index = reind
        else:
            print("Back to Menu...")
            self.clearConsole()
            return
        self.clearConsole()

    def deleteCol(self):
        print("WARNING : You can't recover your data after you delete Row.")
        select = input("Are you sure you choose this option?[y/n] : ")
        colnum = 1
        colsdit = dict()
        if select == 'Y' or select == 'y':
            self.clearConsole()
            print("=" * 10)
            for e in list(self.selectedDB.columns):
                print(colnum,e)
                colsdit[colnum] = e
                colnum += 1
            print("=" * 10)
            print("Enter option number you want to delete")
            selectCol = int(input(">>"))
            if selectCol not in colsdit.keys():
                print("Option Error...")
                self.clearConsole()
                return
            else:
                del self.selectedDB[colsdit[selectCol]]
        else:
            print("Back to Menu...")
            self.clearConsole()
            return
        self.clearConsole()

    def saveData(self):
        select = input("Are you sure you want to save changed data?[y/n] : ")
        if select == 'Y' or select == 'y':
            self.selectedDB.to_excel(dbDir + self.selectedDBName + self.bindedType)
            self.clearConsole()
        else:
            print("Back to Menu...")
            self.clearConsole()

##############################################################
programDir = os.getcwd()
folderList = os.listdir(programDir)

if 'dbSaver' not in folderList:
    os.mkdir(programDir + "/" + 'dbSaver')
else:
    pass

dbDir = programDir + "/" + "dbSaver" + "/"

op = moneyKeep(dbDir)