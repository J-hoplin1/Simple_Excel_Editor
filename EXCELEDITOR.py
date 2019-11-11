import pandas as pd
import os
import time
import sys
import warnings

class excelEditor():
    bindedType = '.xlsx'
    selectedDB = None
    dbDir = None
    selectedDBName = None
    loop1 = 0
    #Index
    selectedDBIndex = None
    #Columns
    selectedDBCol = None

    #for sync DataFrame
    syncData = None

    #현재 지정 DB가 변경사항이 있는지에 대해 체크
    selectedDBStatusChange = False
    selectedDBDataDict = dict()
    '''
       데이터베이스 형성 및 작동 순서 메뉴얼
       1 . 인덱스 추가시 selectedDBIndex변수에 저장
       2 . 값추가시 selectedDBCOL
       3 . 값 추가 혹은 편집시 selectedDBStatusChange변수의 Boolean타입 변경
       4 . menu메소드 안에 main loop가 있음. 순서는 menu -> menuSelect 순을 따른다.
    '''

    #콘솔창 클리어
    def __init__(self, dbdir):
        self.dbDir = dbdir
        self.dbSelect()

    @staticmethod
    def findMeasure(self,dbLength):
        leastNum = 2
        loopBoolean = True
        while loopBoolean:
            if dbLength % leastNum == 0:
                return leastNum
            else:
                leastNum += 1
                continue

    def clearConsole(self):
        os.system('cls')
        time.sleep(0.5)

    def syncDataFrame(self):
        newDataFrame  = pd.DataFrame(
            columns=self.selectedDBCol,
            index=self.selectedDBIndex,
            data=self.selectedDBDataDict
        )
        self.selectedDB = newDataFrame

    def dbSelect(self):
        loopBoolean1 = True
        dbdict = dict()
        while loopBoolean1:
            dbnum = 1
            print('=' * 10 + 'Select Database you want to handle' + '=' * 10)
            print('-1 . Exit')
            print('0 . Create New DB')
            dbli = os.listdir(self.dbDir)
            dbli = [file for file in dbli if file.endswith(".xlsx")]
            for dbname in dbli:
                dbdict[dbnum] = dbname
                print(dbnum, ".", dbname)
                dbnum += 1
            print('=' * 54)
            dbselect = int(input(">>"))
            if dbselect == -1:
                print("Program will be closed soon...")
                self.clearConsole()
            elif dbselect == 0:
                self.clearConsole()
                self.createNewDataBaseDirect()
            elif dbselect not in dbdict.keys():
                print("Please type right number")
            else:
                loopBoolean1 = False
        self.clearConsole()
        # 선택한 DB
        self.selectedDB = pd.read_excel(dbDir + dbdict[dbselect], index_col=0)
        # 선택한 데이터베이스의 이름을 인스턴스 변수에 저장
        self.selectedDBName = dbdict[dbselect].split('.')[0]
        # 선택한 데이터베이스의 인덱스를 인스턴스 변수에 저장
        self.selectedDBIndex = list(self.selectedDB.index)
        self.selectedDBCol = list(self.selectedDB.columns)
        self.selectedDBDataDict = dict()
        for a in self.selectedDBCol:
            self.selectedDBDataDict[a] = list(self.selectedDB[a])
        self.menu()

    #새로운 데이터 생성하기
    def createNewDataBaseDirect(self):
        loopCheck = True
        try:
            newColList = []
            print("Enter new DataBase's Name")
            dbNameInput = input(">>")
            print("How many Columns do you want to generate?")
            generateCol = int(input(">>"))
            for addCol in range(0,generateCol):
                print("Enter name of Col No.",addCol+1)
                colName = input(">>")
                newColList.append(colName)
            newDataFrame = pd.DataFrame(
                columns=newColList
            )
            #새로운 데이터의 저장
            newDataFrame.to_excel(self.dbDir + dbNameInput + '.xlsx')
            print("Initializing new dataframe " + dbNameInput + "...")
            self.clearConsole()
            #만약 지정된 DB가 None(없을때) 생성한 DB를 selectedDB로 설정한다.
            #Initialize DataBase
            if self.selectedDB == None:
                self.selectedDB = newDataFrame
                self.selectedDBName = dbNameInput
                # 선택한 데이터베이스의 인덱스를 인스턴스 변수에 저장
                self.selectedDBIndex = list(self.selectedDB.index)
                self.selectedDBCol = list(self.selectedDB.columns)
                self.selectedDBDataDict = dict()
                for a in self.selectedDBCol:
                    self.selectedDBDataDict[a] = list(self.selectedDB[a])
                self.menu()
            #만약 지정된 데이터베이스가 존재하는 경우 새로운 DB를 selectedDB로 설정할지 의사를 묻는다.
            else:
                print("Here2")
                while loopCheck:
                    print("Do you want to select this DB as default?[Y / N]")
                    answerSelect = input(">>")
                    if answerSelect == 'Y' or answerSelect == 'y':
                        self.selectedDB = newDataFrame
                        loopCheck = False
                        self.clearConsole()
                    elif answerSelect == 'N' or answerSelect == 'n':
                        loopCheck = False
                        self.clearConsole()
                    else:
                        print("Wrong value. Please answer with y or n")
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
            print("Type Error : Back to Main - B")
            self.clearConsole()
            self.dbSelect()
        except KeyError:
            self.loop1 = 0
            print("Key Error : Back to Main")
            self.clearConsole()
            self.dbSelect()

    #menu 메소드 -> menuselect 메소드
    def menuSelect(self):
        try:
            print("Selected Data Base Name : " + str(self.selectedDBName))
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
            self.menuSelect()
            return
        except TypeError:
            print("Type Error : Back to Main-C")
            self.menuSelect()
            return
        except KeyError:
            print("Key Error : Back to Main")
            self.menuSelect()
            return

    #프로그램 종료
    def exitProgram(self):
        print("Program Close...")
        time.sleep(5)
        sys.exit()

    def menu(self):
        loopBoolean2 = True
        # menu loop here
        while loopBoolean2:
            menuNum = self.menuSelect()
            print(menuNum)
            if menuNum == 0:
                self.exitProgram()
            elif menuNum == 1:
                self.createNewDataBase()
            elif menuNum == 2:
                self.viewEntireData()
            elif menuNum == 3:
                if self.selectedDBStatusChange == True:
                    select = input("Do you want to save changed data?[y/n] : ")
                    if select == 'Y' or select == 'y':
                        self.saveData()
                    else:
                        print("Data will not be saved.")
                        self.clearConsole()
                    self.dbSelect()
            elif menuNum == 4:
                self.addData()
            else:
                print("You entered Wrong option number")
                self.clearConsole()


    #전체 데이터 조회하기
    def viewEntireData(self):
        start = 0
        partitionStandard = self.findMeasure(len(self.selectedDB))
        end = partitionStandard
        while end <= len(self.selectedDB):
            print(self.selectedDB.iloc[start: end])
            start += partitionStandard
            end += partitionStandard
        print('\n')
        print('\n')
        print('\n')
        print("Press any key to exit.")
        os.system('pause')

    def saveData(self):
        print("Saving Data...")
        self.selectedDB.to_excel(dbDir + self.selectedDB + self.bindedType)
        self.clearConsole()


    def addData(self):
        loopCK = True
        loopCount = 0
        while loopCK:
            print("=" * 10)
            print("1 . 열 추가하기")
            print("2 . 행 추가하기")
            print("=" * 10)
            select = int(input(">>"))
            self.clearConsole()
            if select == 1:
                print("추가할 열의 개수 정하기")
                selectRowNum = int(input(">>"))
                if selectRowNum > 0:
                    print('1')
                    while loopCount < selectRowNum:
                        print("새로추가할 열 이름 입력하기(Empty Row", loopCount + 1,")")
                        newcolName = input(">>")
                        # 새로운 이름의 열 추가
                        self.selectedDBCol.append(newcolName)
                        # 새로운 열에대한 열 초기화 : 초기값은 NaN으로 됨
                        self.selectedDBDataDict[newcolName] = list()
                        loopCount += 1
                    loopCK = False
                    # 데이터 싱크해주기
                    self.syncDataFrame()
                    self.clearConsole()
                else:
                    print("You entered wrong value. Return to Main")
                    self.clearConsole()
                    self.menu()
            elif select == 2:
                pass


##############################################################



if __name__ == '__main__':
    programDir = os.getcwd()
    folderList = os.listdir(programDir)

    if 'dbSaver' not in folderList:
        os.mkdir(programDir + "/" + 'dbSaver')
    else:
        pass

    dbDir = programDir + "/" + "dbSaver" + "/"

    op = excelEditor(dbDir)