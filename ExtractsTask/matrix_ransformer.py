import os
import numpy as np
import openpyxl
from openpyxl import load_workbook
import statistics

DEBUG = 1

def main():
    sInputFile = 'exp_data.dat'
    if not os.path.exists(sInputFile):
        print("Given File Doesn't Exist!")
        return -1
    with open(sInputFile, 'r') as f:
        arrDataMatrix = [[float(num.replace(',','.')) for num in line.split('\t')] for line in f]
    nRows = len(arrDataMatrix)
    nCols = len(arrDataMatrix[0])

    arrMidCap = list(np.zeros(shape=nCols, dtype=float))
    arrStCap = list(np.zeros(shape=nCols, dtype=float))
    fExelBook = load_workbook('ExtractsTables.xlsx')

    shSourceSheet = fExelBook.worksheets[0]
    shDataSheet = fExelBook.worksheets[1]

    for i in range(nRows):
        for j in range(nCols):
            fCurNumber = arrDataMatrix[i][j]
            shSourceSheet.cell(i+1, j+1).value=fCurNumber


    arrDataMatrixTransposed = np.transpose(arrDataMatrix)
    for i in range(len(arrDataMatrixTransposed)):
        arrMidCap[i] = sum(list(arrDataMatrixTransposed[i]))/nRows
        arrStCap[i] = statistics.stdev(list(arrDataMatrixTransposed[i]))

    for i in range(nRows):
        for j in range(nCols):
            result_number = (arrDataMatrix[i][j]-arrMidCap[j])/arrStCap[j]
            if DEBUG:
                print(f"Result: ({arrDataMatrix[i][j]} - {arrMidCap[j]})/{arrStCap[j]}={result_number}")
            shDataSheet.cell(i+1, j+1).value=result_number
    fExelBook.save('ExtractsTables.xlsx')














if __name__=='__main__':
    main()