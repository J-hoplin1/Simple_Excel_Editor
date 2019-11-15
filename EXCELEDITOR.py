'''
Code Written by J-Hoplin
Repository of this program : https://github.com/J-hoplin1/Simple_Excel_Editor
Opensource License : GPU_3.0
'''
import pandas as pd
import os
import time
import sys
import warnings
import numpy as np
warnings.filterwarnings(action='ignore')

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

    def syncInfo(self):
        self.selectedDBIndex = list(self.selectedDB.index)
        self.selectedDBCol = list(self.selectedDB.columns)
        self.selectedDBDataDict = dict()
        for a in self.selectedDBCol:
            self.selectedDBDataDict[a] = list(self.selectedDB[a])
    def syncDataFrame(self):
        newDataFrame  = pd.DataFrame(
            columns=self.selectedDBCol,
            index=self.selectedDBIndex,
            data=self.selectedDBDataDict
        )
        self.selectedDB = newDataFrame

    def dbSelect(self):
        try:
            loopBoolean1 = True
            dbdict = dict()
            while loopBoolean1:
                dbnum = 1
                print('=' * 10 + 'Select Database you want to handle' + '=' * 10)
                print('-2 . Merge Data')
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
                if dbselect == -2:
                    self.clearConsole()
                    self.mergeData()
                elif dbselect == -1:
                    self.clearConsole()
                    self.exitProgram()
                elif dbselect == 0:
                    self.clearConsole()
                    self.createNewDataBaseDirect()
                elif dbselect not in dbdict.keys():
                    print("Please type right number")
                else:
                    loopBoolean1 = False
            self.clearConsole()
            # 선택한 DB
            self.selectedDB = pd.read_excel(self.dbDir + dbdict[dbselect], index_col=0)
            # 선택한 데이터베이스의 이름을 인스턴스 변수에 저장
            self.selectedDBName = dbdict[dbselect].split('.')[0]
            # 선택한 데이터베이스의 인덱스를 인스턴스 변수에 저장
            self.selectedDBIndex = list(self.selectedDB.index)
            self.selectedDBCol = list(self.selectedDB.columns)
            self.selectedDBDataDict = dict()
            for a in self.selectedDBCol:
                self.selectedDBDataDict[a] = list(self.selectedDB[a])
            self.menu()
        except IndexError:
            print("Error Occured. You need to select again.")
            self.clearConsole()
            self.dbSelect()
        except ValueError:
            print("Error Occured. You need to select again.")
            self.clearConsole()
            self.dbSelect()
        except TypeError:
            print("Error Occured. You need to select again.")
            self.clearConsole()
            self.dbSelect()
        except KeyError:
            print("Error Occured. You need to select again.")
            self.clearConsole()
            self.dbSelect()
        except PermissionError:
            print("Permission Denied. Please check if selected database has been opend")
            self.clearConsole()
            self.dbSelect()

    #새로운 데이터 생성하기
    def createNewDataBaseDirect(self):
        try:
            loopCheck = True #DB save loop
            newColList = []
            print("Enter new DataBase's Name")
            dbNameInput = input(">>")
            print("How many Columns do you want to generate?")
            generateCol = int(input(">>"))
            for addCol in range(0, generateCol):
                print("Enter name of Col No.", addCol + 1)
                colName = input(">>")
                newColList.append(colName)
            newDataFrame = pd.DataFrame(
                columns=newColList
            )
            # 새로운 데이터의 저장
            newDataFrame.to_excel(self.dbDir + dbNameInput + '.xlsx')
            print("Initializing new dataframe " + dbNameInput + "...")
            self.clearConsole()
            # 만약 지정된 DB가 None(없을때) 생성한 DB를 selectedDB로 설정한다.
            # Initialize DataBase
            if self.selectedDB is None:
                self.selectedDB = newDataFrame
                self.selectedDBName = dbNameInput
                # 선택한 데이터베이스의 인덱스를 인스턴스 변수에 저장
                self.selectedDBIndex = list(self.selectedDB.index)
                self.selectedDBCol = list(self.selectedDB.columns)
                self.selectedDBDataDict = dict()
                for a in self.selectedDBCol:
                    self.selectedDBDataDict[a] = list(self.selectedDB[a])
                self.menu()
            # 만약 지정된 데이터베이스가 존재하는 경우 새로운 DB를 selectedDB로 설정할지 의사를 묻는다.
            else:
                while loopCheck:
                    print("Do you want to select this DB as default?[Y / N]")
                    answerSelect = input(">>")
                    if answerSelect == 'Y' or answerSelect == 'y':
                        self.selectedDB = newDataFrame
                        self.syncInfo()
                        loopCheck = False
                        self.clearConsole()
                    elif answerSelect == 'N' or answerSelect == 'n':
                        loopCheck = False
                        self.clearConsole()
                    else:
                        print("Wrong value. Please answer with y or n")
        except IndexError:
            print("Error Occured.You need to generate again.")
            self.clearConsole()
            return
        except ValueError:
            print("Error Occured.You need to generate again.")
            self.clearConsole()
            return
        except TypeError:
            print("Error Occured.You need to generate again.")
            self.clearConsole()
            return
        except KeyError:
            print("Error Occured.You need to generate again.")
            self.clearConsole()
            return
        except PermissionError:
            print("Permission Denied. Please check if selected database has been opend")
            self.clearConsole()
            return

    #menu 메소드 -> menuselect 메소드
    def menuSelect(self):
        try:
            print("Selected Data Base Name : " + str(self.selectedDBName))
            print(self.selectedDB)
            print("=" * 8 + " Menu " + "=" * 8)
            print("0 . Exit Program.")  # 구현완료
            print("1 . Create New DataBase.")  # 구현완료
            print("2 . View entire selected DB.")  # 구현완료
            print("3 . Select another DB.")  # 구현완료
            print("4 . Add Data.")  # 구현완료
            print("5 . View part of selected DB.")  # 구현완료
            print("6 . Delete Row.")  # 구현완료
            print("7 . Delete Column.")  # 구현완료
            print("8 . Save Data.")  # 구현완료
            print("9 . Edit Value")#구현완료
            print("10 . Merge Data")
            print("=" * 22)
            menuNum = int(input(">>"))
            self.clearConsole()
            return menuNum
        except IndexError:
            print("Error Occured.You need to generate again.")
            self.clearConsole()
            return
        except ValueError:
            print("Error Occured.You need to generate again.")
            self.clearConsole()
            return
        except TypeError:
            print("Error Occured.You need to generate again.")
            self.clearConsole()
            return
        except KeyError:
            print("Error Occured.You need to generate again.")
            self.clearConsole()
            return
        except PermissionError:
            print("Permission Denied. Please check if selected database has been opend")
            self.clearConsole()
            return

    #프로그램 종료
    def exitProgram(self):
        print("Program Close...")
        time.sleep(5)
        sys.exit()

    def menu(self):
        try:
            loopBoolean2 = True
            # menu loop here
            while loopBoolean2:
                menuNum = self.menuSelect()
                if menuNum == 0:
                    self.exitProgram()
                elif menuNum == 1:
                    if self.selectedDBStatusChange == False:
                        self.createNewDataBaseDirect()
                    elif self.selectedDBStatusChange == True:
                        select = input("Do you want to save changed data?[y/n] : ")
                        if select == 'Y' or select == 'y':
                            self.saveData()
                        else:
                            print("Data will not be saved.")
                            self.clearConsole()
                        self.createNewDataBaseDirect()
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
                    else:
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
                elif menuNum == 10:
                    self.mergeData()
                else:
                    print("You entered Wrong option number")
                    self.clearConsole()
        except IndexError:
            print("Error Occured.You need to generate again.")
            self.clearConsole()
            self.menu()
        except ValueError:
            print("Error Occured.You need to generate again.")
            self.clearConsole()
            self.menu()
        except TypeError:
            print("Error Occured.You need to generate again.")
            self.clearConsole()
            self.menu()
        except KeyError:
            print("Error Occured.You need to generate again.")
            self.clearConsole()
            self.menu()
        except PermissionError:
            print("Permission Denied. Please check if selected database has been opend")
            self.clearConsole()
            self.menu()


    #전체 데이터 조회하기
    def viewEntireData(self):
        try:
            start = 0
            partitionStandard = self.findMeasure(self, len(self.selectedDB))
            end = partitionStandard
            while end <= len(self.selectedDB):
                print(self.selectedDB.iloc[start: end])
                start += partitionStandard
                end += partitionStandard
            print('\n')
            print('\n')
            print('\n')
            print("Press enter to exit.")
            os.system('pause')
        except IndexError:
            print("Error Occured.You need to generate again.")
            self.clearConsole()
            return
        except ValueError:
            print("Error Occured.You need to generate again.")
            self.clearConsole()
            return
        except TypeError:
            print("Error Occured.You need to generate again.")
            self.clearConsole()
            return
        except KeyError:
            print("Error Occured.You need to generate again.")
            self.clearConsole()
            return
        except PermissionError:
            print("Permission Denied. Please check if selected database has been opend")
            self.clearConsole()
            return

    def saveData(self):
        try:
            print("Saving Data...")
            self.selectedDB.to_excel(dbDir + self.selectedDBName + self.bindedType)
            self.selectedDBStatusChange = False
            self.clearConsole()
        except IndexError:
            print("Error Occured.You need to generate again.")
            self.clearConsole()
            return
        except ValueError:
            print("Error Occured.You need to generate again.")
            self.clearConsole()
            return
        except TypeError:
            print("Error Occured.You need to generate again.")
            self.clearConsole()
            return
        except KeyError:
            print("Error Occured.You need to generate again.")
            self.clearConsole()
            return
        except PermissionError:
            print("Permission Denied. Please check if selected database has been opend")
            self.clearConsole()
            return


    def addData(self):
        try:
            loopCK = True
            loopCount = 0
            while loopCK:
                print("=" * 10)
                print("0 . Exit")
                print("1 . 열 추가하기")
                print("2 . 행 추가하기")
                print("=" * 10)
                select = int(input(">>"))
                self.clearConsole()
                if select == 0:
                    return
                elif select == 1:
                    print("추가할 열의 개수 정하기")
                    selectColNum = int(input(">>"))
                    if selectColNum > 0:
                        while loopCount < selectColNum:
                            print("새로추가할 열 이름 입력하기(Empty column", loopCount + 1, ")")
                            newcolName = input(">>")
                            # 새로운 이름의 열 추가
                            self.selectedDBCol.append(newcolName)
                            # 새로운 열에대한 열 초기화 : 초기값은 NaN으로 됨
                            initList = []
                            for q in range(0, len(self.selectedDBIndex)):
                                print(self.selectedDBIndex[q], "행의 값 추가하기")
                                addVal = input(">>")
                                initList.append(addVal)
                            self.selectedDBDataDict[newcolName] = initList
                            loopCount += 1
                        loopCK = False
                        # 데이터 싱크해주기
                        self.syncDataFrame()
                        self.clearConsole()
                        self.selectedDBStatusChange = True
                    else:
                        print("You entered wrong value. Return to Main")
                        self.clearConsole()
                        self.menu()
                elif select == 2:
                    print("추가할 행의 개수 정하기")
                    selectRowNum = int(input(">>"))
                    if selectRowNum > 0:
                        while loopCount < selectRowNum:
                            print("새로추가할 행 이름 입력하기(Empty Row", loopCount + 1, ")")
                            newRowName = input(">>")
                            self.selectedDBIndex.append(newRowName)
                            for a in self.selectedDBCol:
                                print(a, "열의 값 입력하기")
                                newVal = input(">>")
                                self.selectedDBDataDict[a].append(newVal)
                            loopCount += 1
                        loopCK = False
                        self.syncDataFrame()
                        self.clearConsole()
                        self.selectedDBStatusChange = True
                    else:
                        print("You entered wrong value. Return to Main")
                        self.clearConsole()
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
        except PermissionError:
            print("Permission Denied. Please check if selected database has been opend")
            self.clearConsole()
            return

    def viewAsPart(self):
        try:
            loopCK = True
            while loopCK:
                print("=" * 10)
                print("현재 지정된 데이터베이스 총 행의 개수 : ", len(self.selectedDB), "개의 행")
                print("현재 지정된 데이터베이스 총 열의 개수 : ", len(self.selectedDBCol), "개의 열")
                print("0 . Exit")
                print("1 . 범위를 설정하여 데이터 조회하기")
                print("2 . 특정 열을 지정하여 데이터 조회하기")
                print("=" * 10)
                selectOption = int(input(">>"))
                self.clearConsole()
                if selectOption == 0:
                    return
                elif selectOption == 1:
                    availableRange = list(range(1, len(self.selectedDB) + 1))
                    startRange = int(input("시작할 행의 번호 입력 >>"))
                    endRange = int(input("마지막 행의 번호 입력 >>"))
                    if startRange not in availableRange or endRange not in availableRange:
                        print("잘못된 범위입력. 최소범위는 ", availableRange[0], "이며 최대 범위는 ", availableRange[-1], "입니다")
                    elif startRange > endRange:
                        print("잘못된 범위입력. 최소범위는 ", availableRange[0], "이며 최대 범위는 ", availableRange[-1], "입니다")
                    else:
                        print(self.selectedDB.iloc[startRange - 1:endRange])
                        print('\n')
                        print('\n')
                        print('\n')
                        print("Press any key to exit.")
                        os.system('pause')
                        loopCK = False
                        self.clearConsole()
                elif selectOption == 2:
                    optionDict = dict()
                    select2List = []
                    print("지정할 열의 개수 입력하기")
                    selectColNum = int(input(">>"))
                    if selectColNum not in list(range(1, len(self.selectedDBCol) + 1)):
                        print("최대 열의 개수는 ", len(self.selectedDBCol), "개입니다. 다시입력하세요")
                    else:
                        for r in range(0, selectColNum):
                            print("=" * 10)
                            for a in range(0, len(self.selectedDBCol)):
                                print(a + 1, ".", self.selectedDBCol[a])
                                optionDict[a + 1] = self.selectedDBCol[a]
                            print("=" * 10)
                            option2 = int(input("조회하고자 하는 행의 옵션번호 입력하기 >>"))
                            if option2 not in optionDict.keys():
                                print("잘못된 입력. 방금 입력된 값은 무시됩니다.")
                            else:
                                select2List.append(optionDict[option2])
                                self.clearConsole()
                        select2List.sort()
                        for y in select2List:
                            print(self.selectedDB[y])
                        print('\n')
                        print('\n')
                        print('\n')
                        print("Press any key to exit.")
                        os.system('pause')
                        loopCK = False
                        self.clearConsole()
                else:
                    print("Wrong option number. Please try again.")
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
        except PermissionError:
            print("Permission Denied. Please check if selected database has been opend")
            self.clearConsole()
            return

    def deleteRow(self):
        try:
            loopCK = True
            loopIn = True
            while loopCK:
                print("WARNING : 삭제 후에는 데이터 복구가 불가능합니다. 계속 진행하시겠습니까?[Y / N]")
                select = input(">>")
                if select == 'Y' or select == 'y':
                    self.clearConsole()
                    while loopIn:
                        print("=" * 10)
                        print("0 . Exit")
                        print("1 . 하나의 행 삭제하기")
                        print("2 . 여러개의 행을 지정하여 삭제하기")
                        print("=" * 10)
                        selectOP = int(input(">>"))
                        self.clearConsole()
                        if selectOP == 0:
                            return
                        elif selectOP == 1:
                            print("현재 존재하는 행의 이름들 : ", ",".join(map(str, self.selectedDBIndex)))
                            print("지우고자하는 행의 이름")
                            selectROWNAME = input(">>")
                            if selectROWNAME not in self.selectedDBIndex:
                                print("존재하지 않는 행의 이름입니다.")
                            else:
                                print("데이터 삭제중...")
                                self.selectedDBStatusChange = True
                                self.selectedDB = self.selectedDB.drop(selectROWNAME)
                                self.syncInfo()
                                self.syncDataFrame()
                                self.clearConsole()
                                loopCK = False
                                loopIn = False
                        elif selectOP == 2:
                            li = []
                            print("현재 존재하는 행의 이름들 : ", ",".join(map(str, self.selectedDBIndex)))
                            print("지정할 행의 개수 입력하기")
                            numR = int(input(">>"))
                            for w in range(0, numR):
                                delR = input("행의 이름 입력하기 >>")
                                if delR not in self.selectedDBIndex:
                                    print("존재하지 않는 행이름입니다. 해당 값은 무시됩니다")
                                else:
                                    li.append(delR)
                            print("데이터 삭제중...")
                            self.selectedDBStatusChange = True
                            self.selectedDB = self.selectedDB.drop(li)
                            self.syncInfo()
                            self.syncDataFrame()
                            self.clearConsole()
                            loopIn = False
                            loopCK = False
                        else:
                            print("잘못된 옵션번호. 다시 입력해주세요")
                            self.clearConsole()
                elif select == 'N' or select == 'N':
                    return
                else:
                    print("잘못된 옵션선택. Y 혹은 N으로 답을 입력해주세요.")
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
        except PermissionError:
            print("Permission Denied. Please check if selected database has been opend")
            self.clearConsole()
            return

    def deleteCol(self):
        try:
            loopCK = True
            loopIn = True
            while loopCK:
                print("WARNING : 삭제 후에는 데이터 복구가 불가능합니다. 계속 진행하시겠습니까?[Y / N]")
                select = input(">>")
                if select == 'Y' or select == 'y':
                    self.clearConsole()
                    while loopIn:
                        print("=" * 10)
                        print("0 . Exit")
                        print("1 . 하나의 행 삭제하기")
                        print("2 . 여러개의 행을 지정하여 삭제하기")
                        print("=" * 10)
                        selectOP = int(input(">>"))
                        self.clearConsole()
                        if selectOP == 0:
                            return
                        elif selectOP == 1:
                            print("현재 존재하는 열의 이름들 : ", ",".join(map(str, self.selectedDBCol)))
                            print("지우고자하는 열의 이름")
                            selectColNAME = input(">>")
                            if selectColNAME not in self.selectedDBCol:
                                print("존재하지 않는 행의 이름입니다.")
                            else:
                                print("데이터 삭제중...")
                                del self.selectedDB[selectColNAME]
                                self.selectedDBStatusChange = True
                                self.syncInfo()
                                self.syncDataFrame()
                                self.clearConsole()
                                loopCK = False
                                loopIn = False
                        elif selectOP == 2:
                            li = []
                            print("현재 존재하는 열의 이름들 : ", ",".join(map(str, self.selectedDBCol)))
                            print("지정할 열의 개수 입력하기")
                            numC = int(input(">>"))
                            for w in range(0, numC):
                                delC = input("열의 이름 입력하기 >>")
                                if delC not in self.selectedDBCol:
                                    print("존재하지 않는 열 이름입니다. 해당 값은 무시됩니다")
                                else:
                                    li.append(delC)
                            print("데이터 삭제중...")
                            for o in li:
                                del self.selectedDB[o]
                            self.selectedDBStatusChange = True
                            self.syncInfo()
                            self.syncDataFrame()
                            self.clearConsole()
                            loopCK = False
                            loopIn = False
                        else:
                            print("잘못된 옵션번호. 다시 입력해주세요")
                            self.clearConsole()
                elif select == 'N' or select == 'N':
                    return
                else:
                    print("잘못된 옵션선택. Y 혹은 N으로 답을 입력해주세요.")
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
        except PermissionError:
            print("Permission Denied. Please check if selected database has been opend")
            self.clearConsole()
            return
    def editValue(self):
        try:
            outerLoop = True
            while outerLoop:
                print("=" * 10)
                print("0 . Exit")
                print("1 . Edit")
                print("=" * 10)
                select9 = int(input(">>"))
                self.clearConsole()
                if select9 == 0:
                    return
                elif select9 == 1:
                    print(self.selectedDB)
                    selectNumCol = int(input("값을 편집할 열의 순서를 입력해주세요 >>"))
                    selectNumRow = int(input("값을 편집할 행의 순서를 입력해주세요 >>"))
                    if selectNumCol < 0 or selectNumCol > len(
                            self.selectedDBCol) or selectNumRow < 0 or selectNumRow > len(self.selectedDBIndex):
                        print("행 혹은 열의 이름을 잘못입력하였습니다. 다시 시도해주세요")
                        self.clearConsole()
                    else:
                        changeVal = input("변경할 값을 입력해주세요 >>")
                        self.selectedDB.iloc[selectNumRow - 1, selectNumCol - 1] = changeVal
                        self.syncInfo()
                        self.syncDataFrame()
                        self.clearConsole()
                        self.selectedDBStatusChange = True
                else:
                    print("잘못된 옵션선택. 다시 시도해 주십시오")
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
        except PermissionError:
            print("Permission Denied. Please check if selected database has been opend")
            self.clearConsole()
            return

    def mergeData(self):
        try:
            outloop = True
            inloop = True
            while outloop:
                if self.selectedDB is None:
                    pass
                else:
                    if self.selectedDBStatusChange == True:
                        print("이 작업을 수행하기 전에는 필수로 변경사항이 저장됩니다.")
                        self.saveData()
                    else:
                        self.clearConsole()
                        pass
                print("=" * 10)
                print("0 . Exit")
                print("1 . 행 방향으로 데이터 합치기")
                print("2 . 열 방향으로 데이터 합치기")
                print("=" * 10)
                select10 = int(input(">>"))
                self.clearConsole()
                if select10 == 0:
                    return
                elif select10 == 1 or select10 == 2:
                    while inloop:
                        dbnum = 1
                        ck = True
                        dbdict = dict()
                        selectli = list()
                        print("=" * 10)
                        dbli = os.listdir(self.dbDir)
                        dbli = [file for file in dbli if file.endswith(".xlsx")]
                        print("합칠 두개의 데이터 선택하기")
                        for dbname in dbli:
                            dbdict[dbnum] = dbname
                            print(dbnum, ".", dbname)
                            dbnum += 1
                        print('=' * 10)
                        for e in range(0, 2):
                            select10a = int(input(">>"))
                            selectli.append(select10a)
                        for a in selectli:
                            if a not in dbdict.keys():
                                ck = False
                            else:
                                continue
                        if ck == False:
                            print("잘못된 옵션선택. 다시 선택해주세요.")
                            self.clearConsole()
                        else:
                            selectdata1 = pd.read_excel(self.dbDir + dbdict[selectli[0]], index_col=0)
                            selectdata2 = pd.read_excel(self.dbDir + dbdict[selectli[1]], index_col=0)
                            newDBName = dbdict[selectli[0]].split('.')[0] + "_" + dbdict[selectli[1]].split('.')[
                                0] + '.xlsx'
                            if select10 == 1:
                                new_data = pd.concat([selectdata1, selectdata2], axis=1)
                                new_data.fillna(np.nan)
                                new_data.to_excel(self.dbDir + newDBName)
                                print("합쳐진 데이터가 저장되었습니다.")
                                self.clearConsole()
                                inloop = False
                                outloop = False
                            else:
                                new_data = pd.concat([selectdata1, selectdata2])
                                new_data.fillna(np.nan)
                                new_data.to_excel(self.dbDir + newDBName)
                                print("합쳐진 데이터가 저장되었습니다.")
                                self.clearConsole()
                                inloop = False
                                outloop = False
                else:
                    print("잘못된 옵션선택. 다시 시도해주세요")
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
        except PermissionError:
            print("Permission Denied. Please check if selected database has been opend")
            self.clearConsole()
            return



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
