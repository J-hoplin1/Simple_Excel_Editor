Simple Excel Editor
===
***
- Simple Excel Editor using pandas library.

- Using Language : Python 3.7

- Dev Env : Jetbrain Pycharm

- Used library & modules : pandas, os,time,sys,warnings,numpy,pyinstaller

- exe file is in directory 'dist'

- ??? : GUI로 안만드는 이유가 뭔가요 : 추후 PyQt5로 GUI화시켜서 올릴 예정입니다
***
- 2019/11/13
    
    - Renew Code : Delete code - User unfriendly
    
    - 20191113.ver exe link : [here](https://drive.google.com/open?id=14cZh265q9hwrnnkNOqq_atpaaAErUyqH)
    
- 2019/11/14

    - Add Function : Able to merge dataset. Limit as only two dataset can be merged
    
    - Raise warnings : Some Warniings were occured. Raise these warnings.
    
    - 20191114.ver exe link : [here](https://drive.google.com/open?id=19QLqjE_e15kC0nMRA05FSXAE0FyX3W0-)
    
- 2019/11/15

    - Bug fix : Infinite Loop in option 10 - 2 (Merge Data - 행방향으로 합치기)
    
    - 20191115.ver exe link : [here](https://drive.google.com/open?id=1TC-oGpYJHAT-Dc17NB0ofLBOhXmmTq6Q)
***
### How to use?


- First when you open program Cli will open database select menu. Two example database were given for test. If you want to use program from first, delete dbSaver folder.

    ![img](ExcelEditorimg/1.PNG)

- When you select Create New DB menu, You can create new DataBase.

    - First you need to enter database's name.

    - Second you need to enter number of columns you want to genereate.

    - According to the number you enter(number of columns), you need to write each columns's name.

        ![img](ExcelEditorimg/2.PNG)

- After you create database, menu will be like this.

    ![img](ExcelEditorimg/3.PNG)

- Let's say i want to add some rows. then select 4 . Add Data. Then menu will appear like this.

    ![img](ExcelEditorimg/4.PNG)

    - When you want to add rows, select option number 2.

    - First you need to write number of row you want to add, and enter data of each columns

        ![img](ExcelEditorimg/5.PNG)
    
    - Then you can see rows are added

        ![img](ExcelEditorimg/6.PNG)

- if you want to edit specified value, select option 9. Menu of option 9 looks like this

    ![img](ExcelEditorimg/7.PNG)

    - if you want to edit select 1

    - In this option you need to enter row and col number each about the data you want to edit

        ![img](ExcelEditorimg/8.PNG)

    - After you go to menu, you can see data has changed.

        ![img](ExcelEditorimg/9.PNG)
    
- If you want to save data select 8

- If you want to change database you want to edit, select option 3

    - If you haven't saved database after edit datas, option that ask you to save will open.

        ![img](ExcelEditorimg/10.PNG)
    
    - After you select another database to edit and go to menu, you can see selected database has been changed.

        ![img](ExcelEditorimg/11.PNG)


- If there's bug in this program, please contact to jhyoon0815103@gmail.com
