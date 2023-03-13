Sub FinancesDataExtracter()
'Export data from bank accounts to excel
'Run this macro and proved path for data and target file

'Constants
Const FirstRowBankAccount = 14
Const MainWorksheetStartRow = 3

Debug.Print ("===== Start macro ======")

'Declare Variables
Dim MainWorkbook As Workbook
Dim MainWorksheet As Worksheet
Dim ExportDataWorkbook As Workbook
Dim LastRowBankAccount As Integer
Dim RangeBankAccount As Range

Dim MainCurrentRow As Integer
Dim ItemExist As Boolean
Dim ItemAdded As Boolean

Dim MainOperation As String
Dim MainOutcome As Variant      'Because it may be "-" when 0
Dim MainIncome As Variant       'Because it may be "-" when 0
Dim MainRemain As Variant       'Because it may be "-" when 0

Dim ExportOperation As String
Dim ExportOutcome As Variant    'Because it may be "-" when 0
Dim ExportIncome As Variant     'Because it may be "-" when 0
Dim ExportRemain As Variant     'Because it may be "-" when 0

Set MainWorkbook = ActiveWorkbook
Set MainWorksheet = MainWorkbook.ActiveSheet

'Open Bank account export data
Workbooks.Open "d:\tmp\BankExports\Bank-Movement.xls"
Set ExportDataWorkbook = ActiveWorkbook
LastRowBankAccount = Cells(Rows.Count, 1).End(xlUp).row

'Set RangeBankAccount = ExportDataWorkbook.ActiveSheet.Range(Cells(FirstRowBankAccount, 1), Cells(LastRowBankAccount, 7))
Set RangeBankAccount = ExportDataWorkbook.ActiveSheet.Range("A14:A16")    'For debug

'Go back to main workbook
MainWorkbook.Activate

For Each rangeRow In RangeBankAccount.Rows
    
    ExportOperation = rangeRow.Columns("C")
    ExportOutcome = rangeRow.Columns("E")
    ExportIncome = rangeRow.Columns("F")
    ExportRemain = rangeRow.Columns("G")
    
    Debug.Print ("Exported Table Date: " & rangeRow.Columns("A"))
    Debug.Print ("Export Outcome-Income: " & ExportOutcome & " - " & ExportIncome)
    
    ItemExist = False
    ItemAdded = False
    MainCurrentRow = MainWorksheetStartRow
    
    'Forward pointer to relevant point expDate < MainDate
    While rangeRow.Columns("A") < MainWorksheet.Cells(MainCurrentRow, "C")
        MainCurrentRow = MainCurrentRow + 1
    Wend
    
    'Pass over all Items with same ExportItem Date
    Do While rangeRow.Columns("A") = MainWorksheet.Cells(MainCurrentRow, "C")
            
        MainOperation = MainWorksheet.Cells(MainCurrentRow, "D")
        MainOutcome = MainWorksheet.Cells(MainCurrentRow, "E")
        MainIncome = MainWorksheet.Cells(MainCurrentRow, "F")
        MainRemain = MainWorksheet.Cells(MainCurrentRow, "G")
        
        Debug.Print ("=== DATE MATCH - " & rangeRow.Columns("A"))
        Debug.Print ("Main Outcome-Income: " & MainOutcome & " - " & MainIncome)
        
        'Test if item already exist in MainTable and skip add part
        If ExportOutcome = MainOutcome And ExportIncome = MainIncome Then
            Debug.Print ("!!! Item already exist - Don't add anything - break while loop")
            ItemExist = True
            Exit Do
        End If
        
        'Date is same but content not - Add new item
        If ExportOutcome <> MainOutcome And ExportIncome <> MainIncome Then
            Debug.Print ("!!! Same Date but item different - ADD IT HERE - row: " & MainCurrentRow)
            ItemAdded = True
        End If
        
        MainCurrentRow = MainCurrentRow + 1
    Loop
    
    'Item have diffirent date - Ensure it not added already and add it
    If Not (ItemExist) And Not (ItemAdded) Then
        Debug.Print ("!!! Date not found - New Item - ADD IT HERE - row: " & MainCurrentRow)
    End If
                    
    Debug.Print ("-")
Next rangeRow

End Sub