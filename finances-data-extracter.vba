Sub FinancesDataExtracter()
    'Export data from bank accounts to excel
    'Run this macro and proved path for data and target file
    
    'Constants
    Const FirstRowBankAccount = 14
    Const MainWorksheetStartRow = 3
    
    Debug.Print ("===== Start macro ======")
    
    'Declare Variables
    
    Dim MainWorksheet       As Worksheet
    Dim ExportDataWorkbook  As Workbook
    Dim LastRowBankAccount  As Integer
    Dim RangeBankAccount    As Range
    
    Dim MainCurrentRow      As Integer
    Dim AddItemRow          As Integer
    Dim ItemExist           As Boolean
    Dim ItemAdded           As Boolean
    Dim AddItem             As Boolean
    
    Dim MainOperation       As String
    Dim MainOutcome         As Variant  'Because it may be "-" when 0
    Dim MainIncome          As Variant  'Because it may be "-" when 0
    Dim MainRemain          As Variant  'Because it may be "-" when 0
    
    Dim ExportOperation     As String
    Dim ExportOutcome       As Variant  'Because it may be "-" when 0
    Dim ExportIncome        As Variant  'Because it may be "-" when 0
    Dim ExportRemain        As Variant  'Because it may be "-" when 0
    
    'Save current workbook range for later when be focused on another worksheet
    Set MainWorkbook = ActiveWorkbook
    Set MainWorksheet = ActiveSheet
    
    'Open Bank account export data
    Workbooks.Open "d:\tmp\BankExports\Bank-Movement.xls"
    Set ExportDataWorkbook = ActiveWorkbook
    LastRowBankAccount = Cells(Rows.Count, 1).End(xlUp).row
    
    Set RangeBankAccount = ExportDataWorkbook.ActiveSheet.Range(Cells(FirstRowBankAccount, 1), Cells(LastRowBankAccount, 7))
    'Set RangeBankAccount = ExportDataWorkbook.ActiveSheet.Range("A14:A20")    'For debug
    
    'Go back to main workbook
    MainWorkbook.Activate
    
    '-----------------------------
    'Loop over Exported Data Items
    '-----------------------------
    For Each rangeRow In RangeBankAccount.Rows
        
        ExportDate = rangeRow.Columns("A")
        ExportOperation = rangeRow.Columns("C")
        ExportOutcome = rangeRow.Columns("E")
        ExportIncome = rangeRow.Columns("F")
        ExportRemain = rangeRow.Columns("G")
        
        Debug.Print ("----------------- Item Start -------------------")
        Debug.Print ("ExpData> Date: " & ExportDate & " Out-In: " & ExportOutcome & " - " & ExportIncome)
        
        ItemExist = False
        ItemAdded = True
        AddItem = True
        MainCurrentRow = MainWorksheetStartRow
        
        '---------------------------------------------------------
        'Loop over Main Workbook items with Date > ExportItem Date
        '---------------------------------------------------------
        While ExportDate < MainWorksheet.Cells(MainCurrentRow, "C")
            MainCurrentRow = MainCurrentRow + 1
        Wend
        
        'Set row for potential add item to main table
        AddItemRow = MainCurrentRow
        
        '--------------------------------------------------------
        'Loop over Main Workbook Items with date = exportItemDate
        '--------------------------------------------------------
        Do While ExportDate = MainWorksheet.Cells(MainCurrentRow, "C")
                
            MainOperation = MainWorksheet.Cells(MainCurrentRow, "D")
            MainOutcome = MainWorksheet.Cells(MainCurrentRow, "E")
            MainIncome = MainWorksheet.Cells(MainCurrentRow, "F")
            MainRemain = MainWorksheet.Cells(MainCurrentRow, "G")
            
            Debug.Print ("=== DATE MATCH on row: " & MainCurrentRow & " Date: " & MainWorksheet.Cells(MainCurrentRow, "C"))
            Debug.Print ("Main Outcome-Income: " & MainOutcome & " - " & MainIncome)
            
            'Test if item already exist in MainTable and skip add part
            If ((ExportOutcome = MainOutcome) And _
                (ExportIncome = MainIncome) And _
                (ExportRemain = MainRemain)) Then
                
                Debug.Print ("!!! Item already found - Set AddItem = False")
                ItemExist = True
                AddItem = False
                Exit Do
                
            End If
            
            MainCurrentRow = MainCurrentRow + 1
            Debug.Print ("Item with same date not found - AddItem = " & AddItem)
        Loop
        
        Debug.Print ("End of Main workbook loop")
        
        'When loop throgh dates found it not exist - Add it
        If (AddItem) Then
            
            Debug.Print ("Row for insert : " & MainCurrentRow)
            ItemAdded = True
            MainWorksheet.Cells(MainCurrentRow, "A").EntireRow.Insert
            MainWorksheet.Cells(MainCurrentRow, "C").Value = rangeRow.Columns("A")
            MainWorksheet.Cells(MainCurrentRow, "D").Value = ExportOperation
            MainWorksheet.Cells(MainCurrentRow, "E").Value = ExportOutcome
            MainWorksheet.Cells(MainCurrentRow, "F").Value = ExportIncome
            MainWorksheet.Cells(MainCurrentRow, "G").Value = ExportRemain
            
            'Move pointer to next row
            'MainCurrentRow = AddItemRow + 1
            Debug.Print ("Row Add finish - Current row to check: " & MainCurrentRow)
            Debug.Print ("MainOutcome " & MainWorksheet.Cells(MainCurrentRow, "E"))
        Else
            Debug.Print ("Don't Add item")
        End If
        
        Debug.Print ("-")
    Next rangeRow

End Sub