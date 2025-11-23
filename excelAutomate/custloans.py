custloanNameCol = 4
custloanAmtCol = custloanNameCol+1

def write_custloan_details(sheet, rowNo, custloans):
    custloansCnt = 0
    totCustLoan = 0
    for custloan in custloans:
        #print(rowNo, custloan)
        custloanAmt = float(custloan['custloanAmt'])

        sheet.cell(rowNo, custloanNameCol).value = custloan['custloanName']
        sheet.cell(rowNo, custloanAmtCol).value = custloanAmt
        rowNo+=1
        custloansCnt+=1

        totCustLoan = totCustLoan + float(custloan['custloanAmt'])

    return rowNo,custloansCnt

