import pandas as pd
import os
import time
import sys
import warnings

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
        try:
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
                    self.clearConsole()

            self.clearConsole()
            # 선택한 DB
            self.loop1 += 1
            self.selectedDB = pd.read_excel(dbDir + dbdict[dbselect], index_col=0)
            self.selectedDBName = dbdict[dbselect].split('.')[0]
            self.menu()
        except IndexError:
            print("Index out of bound")
            self.clearConsole()
            return
        except ValueError:
            print("Value Error : Back to Main")
            self.clearConsole()
            return
        except TypeError:
            print("Type Error : Back to Main")
            self.clearConsole()
            return
        except KeyError:
            print("Key Error : Back to Main")
            self.clearConsole()
            return

    def menu(self):
        try:
            loopBoolean2 = True
            # menu loop here
            while loopBoolean2:
                menuNum = self.menuSelect()
                if menuNum == 0:
                    self.exitProgram()
                elif menuNum == 1:
                    self.createNewDataBase()
                elif menuNum == 2:
                    pass
                elif menuNum == 3:
                    select = input("Do you want to save changed data?[y/n] : ")
                    if select == 'Y' or select == 'y':
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
                elif menuNum == 9:
                    self.editValue()
                else:
                    print("You entered Wrong option number")
                    self.clearConsole()
        except IndexError:
            print("Index out of bound")
            self.clearConsole()
            return
        except ValueError:
            print("Value Error : Back to Main")
            self.clearConsole()
            return
        except TypeError:
            print("Type Error : Back to Main")
            self.clearConsole()
            return
        except KeyError:
            print("Key Error : Back to Main")
            self.clearConsole()
            return


    def menuSelect(self):
        try:
            print("Selected Data Base Name : " + self.selectedDBName)
            print(self.selectedDB)
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
            print("9 . Edit Value")
            print("=" * 22)
            menuNum = int(input(">>"))
            self.clearConsole()
            return menuNum
        except IndexError:
            print("Index out of bound")
            self.clearConsole()
            return
        except ValueError:
            print("Value Error : Back to Main")
            self.clearConsole()
            return
        except TypeError:
            print("Type Error : Back to Main")
            self.clearConsole()
            return
        except KeyError:
            print("Key Error : Back to Main")
            self.clearConsole()
            return

    def exitProgram(self):
        print("Program Close...")
        time.sleep(5)
        sys.exit()

    def createNewDataBase(self):
        try:
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
                self.selectedDB.to_excel(dbDir + self.selectedDBName + self.bindedType)
                self.selectedDBName = dbName
                self.selectedDB = pd.read_excel(dbDir + dbName + '.xlsx', index_col=0)
                self.clearConsole()
            elif changeDBYN == 'N' or changeDBYN == 'n':
                print("DataBase " + dbName + " generated in dbSaver Directory")
                self.clearConsole()
        except IndexError:
            print("Index out of bound")
            self.clearConsole()
            return
        except ValueError:
            print("Value Error : Back to Main")
            self.clearConsole()
            return
        except TypeError:
            print("Type Error : Back to Main")
            self.clearConsole()
            return
        except KeyError:
            print("Key Error : Back to Main")
            self.clearConsole()
            return
    def createNewDataBaseDirect(self):
        try:
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
        except IndexError:
            print("Index out of bound")
            self.loop1 = 0
            self.clearConsole()
            return
        except ValueError:
            print("Value Error : Back to Main")
            self.loop1 = 0
            self.clearConsole()
            self.dbSelect()
        except TypeError:
            self.loop1 = 0
            print("Type Error : Back to Main")
            self.clearConsole()
            self.dbSelect()
        except KeyError:
            self.loop1 = 0
            print("Key Error : Back to Main")
            self.clearConsole()
            self.dbSelect()


    def viewSelectDB(self):
        print(self.selectedDB)

    def selectAntherDB(self):
        self.dbSelect()

    def addData(self):
        try:
            loopBoolean = True
            selectOp = None
            print("=" * 10, " Add Data ", "=" * 10)
            print("Total List of Columns : ", ",".join(map(str, list(self.selectedDB.columns))))
            print("=" * 30)
            print("Enter number of Rows you want to add")
            rowCount = int(input(">>"))
            if rowCount == 'exit':
                self.clearConsole()
                return
            indexList = list(self.selectedDB.index)
            if len(indexList) == 0:
                print("=" * 20)
                print("1 . Set index as default(Number)")
                print("2 . Set index as String")
                print("=" * 20)
                selectOp = int(input(">>"))
            elif type(indexList[-1]) == str:
                selectOp = 2
            else:
                print("=" * 20)
                print("1 . Set index as default(Number)")
                print("2 . Set index as String")
                print("=" * 20)
                selectOp = int(input(">>"))
            nextIndexNumber = 1
            if selectOp == 1:
                if len(indexList) is 0:
                    pass
                else:
                    nextIndexNumber = indexList[-1] + 1
                for q in range(0, rowCount):
                    dataSocket = dict()
                    print("=" * 20)
                    for w in list(self.selectedDB.columns):
                        print("Enter value of " + w)
                        val = input(">>")
                        dataSocket[w] = val
                    df = pd.DataFrame(
                        data=dataSocket,
                        index=[nextIndexNumber]
                    )
                    self.selectedDB = pd.concat([self.selectedDB, df])
                    nextIndexNumber += 1
                    print("=" * 20)
            elif selectOp == 2:
                reind = []
                mode = 2
                for er in range(1, len(self.selectedDB) + 1):
                    reind.append(str(er))
                self.selectedDB.index = reind
                for q in range(0, rowCount):
                    dataSocket = dict()
                    print("=" * 20)
                    print("Enter index name")
                    valInd = input(">>")
                    for w in list(self.selectedDB.columns):
                        print("Enter value of " + w)
                        val = input(">>")
                        dataSocket[w] = val
                    df = pd.DataFrame(
                        data=dataSocket,
                        index=[valInd]
                    )
                    self.selectedDB = pd.concat([self.selectedDB, df])
                    nextIndexNumber += 1
                    print("=" * 20)
            self.clearConsole()
        except IndexError:
            print("Index out of bound")
            self.clearConsole()
            return
        except ValueError:
            print("Value Error : Back to Main")
            self.clearConsole()
            return
        except TypeError:
            print("Type Error : Back to Main")
            self.clearConsole()
            return
        except KeyError:
            print("Key Error : Back to Main")
            self.clearConsole()
            return

    def viewAsPart(self):
        try:
            print("Enter number of rows you want to see as part.")
            partNum = int(input(">>"))
            self.clearConsole()
            li = []
            if type(list(self.selectedDB.index)[0]) == int:
                for q in range(0, partNum):
                    print("Enter row name")
                    indNum = int(input(">>"))
                    if indNum not in list(self.selectedDB.index):
                        print("Row name", indNum, "not exist. Ignore this value.")
                        self.clearConsole()
                        pass
                    else:
                        li.append(indNum)
                        self.clearConsole()
                print("=" * 20)
                for w in li:
                    print(self.selectedDB.loc[w])
                    print("=" * 20)
            elif type(list(self.selectedDB.index)[0]) == str:
                for q in range(0, partNum):
                    print("Enter row Name")
                    indNum = input(">>")
                    if indNum not in list(self.selectedDB.index):
                        print("Row name ", indNum, "not exist. Ignore this value.")
                        self.clearConsole()
                        pass
                    else:
                        li.append(indNum)
                        self.clearConsole()
                print("=" * 20)
                for w in li:
                    print(self.selectedDB.loc[w])
                    print("=" * 20)
        except IndexError:
            print("Index out of bound")
            self.clearConsole()
            return
        except ValueError:
            print("Value Error : Back to Main")
            self.clearConsole()
            return
        except TypeError:
            print("Type Error : Back to Main")
            self.clearConsole()
            return
        except KeyError:
            print("Key Error : Back to Main")
            self.clearConsole()
            return

    def deleteRow(self):
        try:
            print("WARNING : You can't recover your data after you delete Row.")
            select = input("Are you sure you choose this option?[y/n] : ")
            if select == 'Y' or select == 'y':
                self.clearConsole()
                if type(list(self.selectedDB.index)[0]) == int:
                    print("Enter Row number you want to delete. There are total ", len(self.selectedDB), "rows.")
                    selectRow = int(input(">>"))
                    if selectRow - 1 not in range(0, len(self.selectedDB)):
                        print("There are no index number : ", selectRow)
                        self.clearConsole()
                        return
                    else:
                        self.selectedDB = self.selectedDB.drop(selectRow)
                        print("Reindexing...")
                        reind = []
                        for er in range(1, len(self.selectedDB) + 1):
                            reind.append(er)
                        self.selectedDB.index = reind
                elif type(list(self.selectedDB.index)[0]) == str:
                    print("Enter Row name you want to delete.")
                    selectRow = input(">>")
                    if selectRow not in list(self.selectedDB.index):
                        print("There are no index name : ", selectRow)
                        self.clearConsole()
                        return
                    else:
                        self.selectedDB = self.selectedDB.drop(selectRow)
                else:
                    print("Back to Menu...")
                    self.clearConsole()
                    return
            self.clearConsole()
        except IndexError:
            print("Index out of bound")
            self.clearConsole()
            return
        except ValueError:
            print("Value Error : Back to Main")
            self.clearConsole()
            return
        except TypeError:
            print("Type Error : Back to Main")
            self.clearConsole()
            return
        except KeyError:
            print("Key Error : Back to Main")
            self.clearConsole()
            return

    def deleteCol(self):
        try:
            print("WARNING : You can't recover your data after you delete Row.")
            select = input("Are you sure you choose this option?[y/n] : ")
            colnum = 1
            colsdit = dict()
            if select == 'Y' or select == 'y':
                self.clearConsole()
                print("=" * 10)
                for e in list(self.selectedDB.columns):
                    print(colnum, e)
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
        except IndexError:
            print("Index out of bound")
            self.clearConsole()
            return
        except ValueError:
            print("Value Error : Back to Main")
            self.clearConsole()
            return
        except TypeError:
            print("Type Error : Back to Main")
            self.clearConsole()
            return
        except KeyError:
            print("Key Error : Back to Main")
            self.clearConsole()
            return


    def saveData(self):
        try:
            select = input("Are you sure you want to save changed data?[y/n] : ")
            if select == 'Y' or select == 'y':
                self.selectedDB.to_excel(dbDir + self.selectedDBName + self.bindedType)
                self.clearConsole()
            else:
                print("Back to Menu...")
                self.clearConsole()
        except IndexError:
            print("Index out of bound")
            self.clearConsole()
            return
        except ValueError:
            print("Value Error : Back to Main")
            self.clearConsole()
            return
        except TypeError:
            print("Type Error : Back to Main")
            self.clearConsole()
            return
        except KeyError:
            print("Key Error : Back to Main")
            self.clearConsole()
            return



    def editValue(self):
        try:
            mode = 0
            print("Select Row index you want to edit(Enter exit to back to menu)")
            select = input(">>")
            if select == "exit":
                return
            if type(list(self.selectedDB.index)[0]) == int:
                mode = 1
                select = int(select)
                if select not in list(self.selectedDB.index):
                    print("That index does not exist.")
                    self.clearConsole()
                    return
            elif type(list(self.selectedDB.index)[0]) == str:
                select = str(select)
                mode = 2
                if select not in list(self.selectedDB.index):
                    print("That index does not exist.")
                    self.clearConsole()
                    return
            elif type(list(self.selectedDB.index)[0]) == str and type((self.selectedDB.index)[0]) == int:
                reind = []
                mode = 2
                for er in range(1, len(self.selectedDB) + 1):
                    reind.append(str(er))
                self.selectedDB.index = reind
            self.clearConsole()
            loopBoolean = True
            while loopBoolean:
                colnum = 1
                colsdit = dict()
                print("=" * 10)
                print('0 . Exit')
                for e in list(self.selectedDB.columns):
                    if mode == 1:
                        print(colnum, '.', e, ':', self.selectedDB[e][select])
                    elif mode == 2:
                        print(colnum,'.',e,':',self.selectedDB[e][select])
                    colsdit[colnum] = e
                    colnum += 1
                print("=" * 10)
                selectOption = int(input("Select option number you want to edit : "))
                if selectOption == 0:
                    print("Go Back to Menu")
                    self.clearConsole()
                    loopBoolean = False
                elif selectOption not in colsdit.keys():
                    print("Option you select didn't exist")
                    continue
                else:
                    changeVal = input("Enter Value you want to change : ")
                    self.selectedDB[colsdit[selectOption]][select] = changeVal
                    warnings.filterwarnings(action='ignore')
                    self.clearConsole()
                self.clearConsole()
        except IndexError:
            print("Index out of bound")
            self.clearConsole()
            return
        except ValueError:
            print("Value Error : Back to Main")
            self.clearConsole()
            return
        except TypeError:
            print("Type Error : Back to Main")
            self.clearConsole()
            return
        except KeyError:
            print("Key Error : Back to Main")
            self.clearConsole()
            return



##############################################################
programDir = os.getcwd()
folderList = os.listdir(programDir)

if 'dbSaver' not in folderList:
    os.mkdir(programDir + "/" + 'dbSaver')
else:
    pass

dbDir = programDir + "/" + "dbSaver" + "/"

op = moneyKeep(dbDir)