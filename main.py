import random
import sys

import numpy as np
import matplotlib.pyplot as plt
from scipy.stats import t
from scipy.stats import uniform
import xlwings as XW



def Rvs(npArray, Df, Size, Min, Max):
    n_Mean = np.mean(npArray)
    n_Std  = np.std(npArray)

    Result = t.rvs(Df, loc=n_Mean, scale=n_Std, size=Size).tolist()
    while True:
        Result = [ x for x in Result if x > Min and x < Max ]
        if(len(Result) == Size):
            break

        Data = t.rvs(Df, loc=n_Mean, scale=n_Std, size=Size-len(Result)).tolist()
        Result = Result + Data
        random.shuffle(Result)

    return Result


def RoundList(Array):
    Array = [ round(x, 1) for x in Array]
    return Array



def Draw_t(n_Mean, n_Std, df):
    Data = np.random.standard_t(df, size=1000) * np.sqrt(n_Std) + n_Mean

    # 绘制直方图
    plt.hist(Data, bins=50, density=True)
    plt.xlabel('Data')
    plt.ylabel('Frequency')
    plt.title('Histogram of t-distribution')
    plt.show()



def ReadExcelCol(Sheet, Col, RowStart, RowEnd):
    Result = []

    for Row in range(RowStart, RowEnd + 1):
        Value = Sheet.range(Col + str(Row)).value
        if Value != None:
            Result.append(Value)

    return Result



def ReadExcelColRate(Sheet, UpperCol, DownCol, RowStart, RowEnd):
    Result = []

    for Row in range(RowStart, RowEnd + 1):
        UpperValue = Sheet.range(UpperCol + str(Row)).value
        DownValue = Sheet.range(DownCol + str(Row)).value
        if UpperValue != None and DownValue != None :
            Result.append( UpperValue/DownValue )

    return Result





def WriteExcelCol(Sheet, Col, RowStart, RowEnd, Values):
    Index = 0
    for Row in range(RowStart, RowEnd + 1):
        Value = Sheet.range(Col + str(Row)).value
        if Value == None:
            Sheet.range(Col + str(Row)).value = Values[Index]
            Index += 1

def WriteExcelColRate(Sheet, UpperCol, DownCol, RowStart, RowEnd, Values):
    Index = 0
    for Row in range(RowStart, RowEnd + 1):
        UpperValue = Sheet.range(UpperCol + str(Row)).value
        DownValue = Sheet.range(DownCol + str(Row)).value
        if UpperValue == None and DownValue != None:
            Sheet.range(UpperCol + str(Row)).value = round(Values[Index] * DownValue, 1)
            Index += 1



def WriteSameData(Sheet, Col, Min, Max):
    Data = np.array(ReadExcelCol(Sheet, Col, 2, 65))
    RandomData = Rvs(Data, 63, 64 - len(Data), Min, Max)
    WriteExcelCol(Sheet, Col, 2, 65, RoundList(RandomData))
    Book.save()

def WriteSameDataRate(Sheet, UpperCol, DownCol, Min, Max):
    Data = np.array(ReadExcelColRate(Sheet, UpperCol, DownCol, 2, 65))
    RandomData = Rvs(Data, 63, 64 - len(Data), Min, Max)
    WriteExcelColRate(Sheet, UpperCol, DownCol, 2, 65, RandomData)
    Book.save()




if __name__ == '__main__':
    App = XW.App(visible=True, add_book=False)
    Book = App.books.open(r'C:\Users\FlyFy\Documents\EMRAssistant2\LisReportData\Fyong.xlsx')
    # 打开第一章表
    Sheet = Book.sheets('Sheet1')

    #第一天总胆红素
    WriteSameData(Sheet, 'K', 500, 50000)


    #第一天结合胆红素
    WriteSameDataRate(Sheet, 'L', 'K', 0.15, 0.99)


    #第五天总胆红素
    Data = np.array(ReadExcelColRate(Sheet, 'N', 'K', 2, 65))
    RandomData = Rvs(Data, 31, 32-3, 0.25, 0.75)
    WriteExcelColRate(Sheet, 'N', 'K', 2, 33, RandomData)

    RandomData = Rvs(Data, 31, 32 - 3, 0.5, 1.2)
    WriteExcelColRate(Sheet, 'N', 'K', 34, 65, RandomData)
    Book.save()

    # 第五天非结合胆红素
    Data = np.array(ReadExcelColRate(Sheet, 'O', 'N', 2, 65))
    RandomData = Rvs(Data, 31, 32 - 3, 0.1, 0.4)
    WriteExcelColRate(Sheet, 'O', 'N', 2, 33, RandomData)

    RandomData = Rvs(Data, 31, 32 - 3, 0.2, 0.5)
    WriteExcelColRate(Sheet, 'O', 'N', 34, 65, RandomData)

    #血脂肪酶
    WriteSameData(Sheet, 'Z', 10, 10000)
    WriteSameDataRate(Sheet, 'AA', 'Z', 0.01, 3)

    #血淀粉酶
    WriteSameData(Sheet, 'AB', 10, 10000)
    WriteSameDataRate(Sheet, 'AC', 'AB', 0.01, 3)

    #尿淀粉酶
    WriteSameData(Sheet, 'AD', 10, 10000)
    WriteSameDataRate(Sheet, 'AE', 'AD', 0.01, 3)

    #天冬
    WriteSameData(Sheet, 'AF', 10, 1000)
    WriteSameDataRate(Sheet, 'AG', 'AF', 0.5, 3)

    #丙氨酸
    WriteSameData(Sheet, 'AH', 10, 1000)
    WriteSameDataRate(Sheet, 'AI', 'AH', 0.5, 3)

    #谷氨酸
    WriteSameData(Sheet, 'AJ', 10, 1000)
    WriteSameDataRate(Sheet, 'AK', 'AJ', 0.5, 3)

    #ALP
    WriteSameData(Sheet, 'AL', 10, 1000)
    WriteSameDataRate(Sheet, 'AM', 'AL', 0.5, 3)

    #总胆红素
    WriteSameData(Sheet, 'AN', 10, 1000)
    WriteSameDataRate(Sheet, 'AO', 'AN', 0.5, 3)

    #结合胆红素
    WriteSameDataRate(Sheet, 'AQ', 'AO', 0.01, 0.5)

    #非结合胆红素
    WriteSameData(Sheet, 'AR', 1, 50)
    WriteSameDataRate(Sheet, 'AS', 'AR', 0.25, 1.2)

    i = 0


# See PyCharm help at https://www.jetbrains.com/help/pycharm/
