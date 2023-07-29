Sub FinancesDataExtracter()
    'Export data from bank accounts to excel
    'Run this macro and proved path for data and target file
    
    'Constants
    Const SourceFirstRowOfTable = 14
    Const TargetSheetStartRow = 3
    Const TargetColumnDate = "C"
    Const TargetColumnOperationName = "D"
    Const TargetColumnSpent = "E"
    Const mColumnIncome = "F"
    Const mColumnRemain = "G"
    
    Debug.Print ("===== Start macro ======")
    
    'Declare Variables
    
    Dim TargetWorksheet       As Worksheet
    Dim SourceDataWorkbook  As Workbook
    Dim LastRowBankAccount  As Integer
    Dim RangeBankAccount    As Range
    
    Dim MainCurrentRow      As Integer
    Dim AddItemRow          As Integer
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
    Set TargetWorksheet = ActiveSheet
    
    'Open Bank account export data
    Workbooks.Open "d:\tmp\BankExports\Bank-Movement.xls"
    Set SourceDataWorkbook = ActiveWorkbook
    LastRowBankAccount = Cells(Rows.Count, 1).End(xlUp).row
    
    Set RangeBankAccount = SourceDataWorkbook.ActiveSheet.Range(Cells(SourceFirstRowOfTable, 1), Cells(LastRowBankAccount, 7))
    'Set RangeBankAccount = SourceDataWorkbook.ActiveSheet.Range("A14:A20")    'For debug
    
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
        
        ItemAdded = True
        AddItem = True
        MainCurrentRow = TargetSheetStartRow
        
        '---------------------------------------------------------
        'Loop over Main Workbook items with Date > ExportItem Date
        '---------------------------------------------------------
        While ExportDate < TargetWorksheet.Cells(MainCurrentRow, "C")
            MainCurrentRow = MainCurrentRow + 1
        Wend
        
        'Set row for potential add item to main table
        AddItemRow = MainCurrentRow
        
        '--------------------------------------------------------
        'Loop over Main Workbook Items with date = exportItemDate
        '--------------------------------------------------------
        Do While ExportDate = TargetWorksheet.Cells(MainCurrentRow, TargetColumnDate)
                
            MainOperation = TargetWorksheet.Cells(MainCurrentRow, TargetColumnOperationName)
            MainOutcome = TargetWorksheet.Cells(MainCurrentRow, TargetColumnSpent)
            MainIncome = TargetWorksheet.Cells(MainCurrentRow, mColumnIncome)
            MainRemain = TargetWorksheet.Cells(MainCurrentRow, mColumnRemain)
            
            Debug.Print ("=== DATE MATCH on row: " & MainCurrentRow & " Date: " & TargetWorksheet.Cells(MainCurrentRow, TargetColumnDate))
            Debug.Print ("Main Outcome-Income: " & MainOutcome & " - " & MainIncome)
            
            'Test if item already exist in MainTable and skip add part
            If ((ExportOutcome = MainOutcome) And _
                (ExportIncome = MainIncome) And _
                (ExportRemain = MainRemain)) Then
                
                Debug.Print ("!!! Item already found - Set AddItem = False")
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
            TargetWorksheet.Cells(MainCurrentRow, "A").EntireRow.Insert
            
            'Set values for new empty row
            TargetWorksheet.Cells(MainCurrentRow, TargetColumnDate).value = ExportDate
            TargetWorksheet.Cells(MainCurrentRow, TargetColumnOperationName).value = ExportOperation
            TargetWorksheet.Cells(MainCurrentRow, TargetColumnSpent).value = ExportOutcome
            TargetWorksheet.Cells(MainCurrentRow, mColumnIncome).value = ExportIncome
            TargetWorksheet.Cells(MainCurrentRow, mColumnRemain).value = ExportRemain
            
            Debug.Print ("Row Add finish - Current row to check: " & MainCurrentRow)
        End If
        
        Debug.Print ("-")
    Next rangeRow

End Sub

Sub CreditCardDataExtracter()
    'Export data from Credit card report
    'Run this macro and proved path for data and target file
    
    'Constants
    Const SourceFirstRowOfTable = 12
    Const TargetSheetStartRow = 3
    Const TargetColumnDate = "C"
    Const TargetColumnOperationName = "D"
    Const TargetColumnSpent = "E"
    Const SourceColumnDate = "A"
    Const SourceColumnOperationName = "B"
    Const SourceColumnSpent = "F"
    
    Debug.Print ("===== Start macro ======")
    
    'Declare Variables
    
    Dim TargetWorksheet     As Worksheet
    Dim SourceDataWorkbook  As Workbook
    Dim LastRowBankAccount  As Integer
    Dim RangeBankAccount    As Range
    
    Dim MainCurrentRow      As Integer
    Dim AddItemRow          As Integer
    Dim ItemAdded           As Boolean
    Dim AddItem             As Boolean
    
    Dim MainOperation       As String
    Dim MainOutcome         As Variant  'Because it may be "-" when 0
    
    Dim ExportOperation     As String
    Dim ExportOutcome       As Variant  'Because it may be "-" when 0
    
    'Save current workbook range for later when be focused on another worksheet
    Set MainWorkbook = ActiveWorkbook
    Set TargetWorksheet = ActiveSheet
    
    'Open Bank account export data
    Workbooks.Open "d:\tmp\BankExports\Credit-card.xls"
    Set SourceDataWorkbook = ActiveWorkbook
    LastRowBankAccount = Cells(Rows.Count, 1).End(xlUp).row
    
    Set RangeBankAccount = SourceDataWorkbook.ActiveSheet.Range(Cells(SourceFirstRowOfTable, 1), Cells(LastRowBankAccount, 7))
    'Set RangeBankAccount = SourceDataWorkbook.ActiveSheet.Range("A12:A16")    'For debug
    
    'Go back to main workbook
    MainWorkbook.Activate
    
    '-----------------------------
    'Loop over Exported Data Items
    '-----------------------------
    For Each rangeRow In RangeBankAccount.Rows
        
        ExportDate = rangeRow.Columns(SourceColumnDate)
        ExportOperation = rangeRow.Columns(SourceColumnOperationName)
        ExportOutcome = rangeRow.Columns(SourceColumnSpent)
        
        Debug.Print ("----------------- Item Start -------------------")
        Debug.Print ("ExpData> Date: " & ExportDate & " Spent: " & ExportOutcome)
        
        ItemAdded = True
        AddItem = True
        MainCurrentRow = TargetSheetStartRow
        
        '---------------------------------------------------------
        'Loop over Main Workbook items with Date > ExportItem Date
        '---------------------------------------------------------
        While ExportDate < TargetWorksheet.Cells(MainCurrentRow, "C")
            MainCurrentRow = MainCurrentRow + 1
        Wend
        
        'Set row for potential add item to main table
        AddItemRow = MainCurrentRow
        
        '--------------------------------------------------------
        'Loop over Main Workbook Items with date = exportItemDate
        '--------------------------------------------------------
        Do While ExportDate = TargetWorksheet.Cells(MainCurrentRow, TargetColumnDate)
                
            MainOperation = TargetWorksheet.Cells(MainCurrentRow, TargetColumnOperationName)
            MainOutcome = TargetWorksheet.Cells(MainCurrentRow, TargetColumnSpent)
            
            Debug.Print ("=== DATE MATCH on row: " & MainCurrentRow & " Date: " & TargetWorksheet.Cells(MainCurrentRow, TargetColumnDate))
            Debug.Print ("Main Spent: " & MainOutcome)
            
            'Test if item already exist in MainTable and skip add part
            If (ExportOutcome = MainOutcome) Then
                
                Debug.Print ("!!! Item already found - Set AddItem = False")
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
            TargetWorksheet.Cells(MainCurrentRow, "A").EntireRow.Insert CopyOrigin:=xlFormatFromLeftOrAbove
            
            'Set values for new empty row
            TargetWorksheet.Cells(MainCurrentRow, TargetColumnDate).value = ExportDate
            TargetWorksheet.Cells(MainCurrentRow, TargetColumnOperationName).value = ExportOperation
            TargetWorksheet.Cells(MainCurrentRow, TargetColumnSpent).value = ExportOutcome
            
            Debug.Print ("Row Add finish - Current row to check: " & MainCurrentRow)
        End If
        
        Debug.Print ("-")
    Next rangeRow

End Sub