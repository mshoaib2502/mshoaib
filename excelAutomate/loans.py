from common_functions import getCellNo, next_month

loanMonthCol = 1
loanTotalEMICol = loanMonthCol+1
loanRemAmtCol = loanTotalEMICol+1
loanTotExpCol = loanRemAmtCol+1

def write_loan_details(sheet, rowNo, loans, bankDet, dateDet):

    maxEMI = 0
    totLoan = 0
    totPending = 0
    for loan in loans:
        #print(rowNo, loan)
        loanAmt = float(loan['loanAmt'])
        loanDate = int(loan['loanDate'])

        sheet.cell(rowNo, 1).value = loanDate
        sheet.cell(rowNo, 2).value = loan['loanName']
        sheet.cell(rowNo, 3).value = loanAmt
        sheet.cell(rowNo, 4).value = int(loan['loanEMI'])
        pending = loanAmt * int(loan['loanEMI'])
        pendingFr = "=" + sheet.cell(rowNo, 3).coordinate + " * " + sheet.cell(rowNo, 4).coordinate
        sheet.cell(rowNo, 5).value = pendingFr
        rowNo+=1

        totLoan = totLoan + float(loan['loanAmt'])
        totPending = totPending + pending
        if bankDet.get(loan['loanBank']) == None:
            bankDet[loan['loanBank']] = loanAmt
        else:
            bankDet[loan['loanBank']] = bankDet.get(loan['loanBank']) + loanAmt

        if dateDet.get(loanDate) == None:
            dateDet[loanDate] = loanAmt
        else:
            dateDet[loanDate] = dateDet.get(loanDate) + loanAmt

        if int(loan['loanEMI']) > maxEMI:
            maxEMI = int(loan['loanEMI'])
        
    return rowNo, bankDet, dateDet, totLoan, maxEMI, totPending

def write_loan_emi_details(sheet, currDet, loans, maxEMI, totExpAbsCell, totInvestAbsCell, RemainingLoan, totalLoanAndOthers, rowNo):

    currDate = currDet['date']

    if maxEMI > 0:
        nextMonth = currDate
        sheet.cell(rowNo, loanMonthCol).value = "Month"
        sheet.cell(rowNo, loanTotalEMICol).value = "Total EMI"
        sheet.cell(rowNo, loanRemAmtCol).value = "Remaining Amount"
        sheet.cell(rowNo, loanTotExpCol).value = "Total exp"
        rowNo+=1

        sheet.cell(rowNo, loanRemAmtCol).value = RemainingLoan
        RemainingLoanCell = getCellNo(sheet, rowNo, loanRemAmtCol)
        rowNo+=1

        for i in range(0, maxEMI+1):
            totalEMI = 0
            for loan in loans:
                #print(rowNo, loan)
                loanAmt = float(loan['loanAmt'])
                loanEMI = int(loan['loanEMI'])
                if loanEMI - i > 0:
                    totalEMI = totalEMI + loanAmt

            if totalEMI > 0:
                nextMonth = next_month(nextMonth)
                RemainingLoan = RemainingLoan - totalEMI
                totalLoanAndEMI = totalLoanAndOthers + totalEMI

                sheet.cell(rowNo, loanMonthCol).value = nextMonth
                sheet.cell(rowNo, loanTotalEMICol).value = totalEMI
                sheet.cell(rowNo, loanRemAmtCol).value = "=" + getCellNo(sheet, rowNo-1, loanRemAmtCol) +"-" + getCellNo(sheet, rowNo, loanTotalEMICol)
                sheet.cell(rowNo, loanTotExpCol).value = "=" + totExpAbsCell + "+" + totInvestAbsCell + "+" + str(totalEMI)
                rowNo+=1
