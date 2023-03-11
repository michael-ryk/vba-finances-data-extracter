Sub FinancesDataExtracter()
'Export data from bank accounts to excel
'Run this macro and proved path for data and target file

Const FirstRowBankAccount = 14

Debug.Print ("Start macro")

'Declare
Dim MainWorkbook As Workbook
Dim ExportDataWorkbook As Workbook
Dim LastRowBankAccount As Integer
Dim RangeBankAccount As Range

Set MainWorkbook = ActiveWorkbook

'Open Bank account expert data
Workbooks.Open "d:\tmp\BankExports\Bank-Movement.xls"
Set ExportDataWorkbook = ActiveWorkbook
LastRowBankAccount = Cells(Rows.Count, 1).End(xlUp).row
Set RangeBankAccount = Range(Cells(FirstRowBankAccount, 1), Cells(LastRowBankAccount, 7))
'Set RangeBankAccount = Range("A14:A20")    'For debug

For Each rangeRow In RangeBankAccount.Rows
    Debug.Print (rangeRow.Address)
Next rangeRow

'Go back to main workbook
MainWorkbook.Activate


End Sub