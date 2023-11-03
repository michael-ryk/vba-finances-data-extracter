Sub FinancesDataExtracter()
    '=======================================================
    'Export data from bank accounts to excel
    'Run this macro and proved path for data and target file
    '=======================================================
    
    'Constants
    Const SourceFirstRowOfTable = 14
    Const TargetSheetStartRow = 3
    Const TargetColumnDate = "C"
    Const TargetColumnOperationName = "D"
    Const TargetColumnSpent = "E"
    Const mColumnIncome = "F"
    Const mColumnRemain = "G"
    
    Debug.Print ("===== Start macro ======")
    
    '------------------
    'Declare Variables
    '------------------
    
    Dim TargetWorksheet       As Worksheet
    Dim SourceDataWorkbook  As Workbook
    Dim SourceLastRow  As Integer
    Dim SourceTableRange    As Range
    
    Dim TargetCurrentRow      As Integer
    Dim AddItemRow          As Integer
    Dim ItemAdded           As Boolean
    Dim AddItem             As Boolean
    
    Dim TargetOperation       As String
    Dim TargetSpent         As Variant  'Because it may be "-" when 0
    Dim MainIncome          As Variant  'Because it may be "-" when 0
    Dim MainRemain          As Variant  'Because it may be "-" when 0
    
    Dim SourceOperation     As String
    Dim SourceSpent       As Variant  'Because it may be "-" when 0
    Dim ExportIncome        As Variant  'Because it may be "-" when 0
    Dim ExportRemain        As Variant  'Because it may be "-" when 0
    
    'Save current workbook range for later when be focused on another worksheet
    Set TargetWorkbook = ActiveWorkbook
    Set TargetWorksheet = ActiveSheet
    
    'Open Bank account export data
    Workbooks.Open "d:\tmp\BankExports\Bank-Movement.xls"
    Set SourceDataWorkbook = ActiveWorkbook
    SourceLastRow = Cells(Rows.Count, 1).End(xlUp).row
    
    Set SourceTableRange = SourceDataWorkbook.ActiveSheet.Range(Cells(SourceFirstRowOfTable, 1), Cells(SourceLastRow, 7))
    'Set SourceTableRange = SourceDataWorkbook.ActiveSheet.Range("A14:A20")    'For debug
    
    'Go back to main workbook
    TargetWorkbook.Activate
    
    '-----------------------------
    'Loop over Exported Data Items
    '-----------------------------
    For Each rangeRow In SourceTableRange.Rows
        
        SourceDate = rangeRow.Columns("A")
        SourceOperation = rangeRow.Columns("C")
        SourceSpent = rangeRow.Columns("E")
        ExportIncome = rangeRow.Columns("F")
        ExportRemain = rangeRow.Columns("G")
        
        Debug.Print ("----------------- Item Start -------------------")
        Debug.Print ("ExpData> Date: " & SourceDate & " Out-In: " & SourceSpent & " - " & ExportIncome)
        
        ItemAdded = True
        AddItem = True
        TargetCurrentRow = TargetSheetStartRow
        
        '---------------------------------------------------------
        'Loop over Main Workbook items with Date > ExportItem Date
        '---------------------------------------------------------
        While SourceDate < TargetWorksheet.Cells(TargetCurrentRow, "C")
            TargetCurrentRow = TargetCurrentRow + 1
        Wend
        
        'Set row for potential add item to main table
        AddItemRow = TargetCurrentRow
        
        '--------------------------------------------------------
        'Loop over Main Workbook Items with date = exportItemDate
        '--------------------------------------------------------
        Do While SourceDate = TargetWorksheet.Cells(TargetCurrentRow, TargetColumnDate)
                
            TargetOperation = TargetWorksheet.Cells(TargetCurrentRow, TargetColumnOperationName)
            TargetSpent = TargetWorksheet.Cells(TargetCurrentRow, TargetColumnSpent)
            MainIncome = TargetWorksheet.Cells(TargetCurrentRow, mColumnIncome)
            MainRemain = TargetWorksheet.Cells(TargetCurrentRow, mColumnRemain)
            
            Debug.Print ("=== DATE MATCH on row: " & TargetCurrentRow & " Date: " & TargetWorksheet.Cells(TargetCurrentRow, TargetColumnDate))
            Debug.Print ("Main Outcome-Income: " & TargetSpent & " - " & MainIncome)
            
            'Test if item already exist in MainTable and skip add part
            If ((SourceSpent = TargetSpent) And _
                (ExportIncome = MainIncome) And _
                (ExportRemain = MainRemain)) Then
                
                Debug.Print ("!!! Item already found - Set AddItem = False")
                AddItem = False
                Exit Do
                
            End If
            
            TargetCurrentRow = TargetCurrentRow + 1
            Debug.Print ("Item with same date not found - AddItem = " & AddItem)
        Loop
        
        Debug.Print ("End of Main workbook loop")
        
        '--------------------------------------------------
        'If Item not exist - Add it
        '--------------------------------------------------
        If (AddItem) Then
            
            Debug.Print ("Row for insert : " & TargetCurrentRow)
            ItemAdded = True
            TargetWorksheet.Cells(TargetCurrentRow, "A").EntireRow.Insert
            
            'Set values for new empty row
            TargetWorksheet.Cells(TargetCurrentRow, TargetColumnDate).Value = SourceDate
            TargetWorksheet.Cells(TargetCurrentRow, TargetColumnOperationName).Value = SourceOperation
            TargetWorksheet.Cells(TargetCurrentRow, TargetColumnSpent).Value = SourceSpent
            TargetWorksheet.Cells(TargetCurrentRow, mColumnIncome).Value = ExportIncome
            TargetWorksheet.Cells(TargetCurrentRow, mColumnRemain).Value = ExportRemain
            
            Debug.Print ("Row Add finish - Current row to check: " & TargetCurrentRow)
        End If
        
        Debug.Print ("-")
    Next rangeRow
    
    SourceDataWorkbook.Close SaveChanges:=False

End Sub

Sub CreditCardDataExtracter()
    '====================================
    'Export data from Credit card report
    'Run this macro and proved path for data and target file
    '====================================
    
    'Constants
    Const SourceFirstRowOfTable = 12
    Const TargetSheetStartRow = 3
    Const TargetColumnDate = "C"
    Const TargetColumnCreditMonth = "D"
    Const TargetColumnOperationName = "E"
    Const TargetColumnSpent = "F"
    Const SourceColumnDate = "A"
    Const SourceColumnOperationName = "B"
    Const SourceColumnSpent = "F"
    
    Debug.Print ("===== Start macro ======")
    
    'Declare Variables
    Dim TargetWorksheet     As Worksheet
    Dim SourceDataWorkbook  As Workbook
    Dim SourceLastRow       As Integer
    Dim SourceTableRange    As Range
    Dim TargetOperation     As String
    Dim TargetSpent         As Variant  'Because it may be "-" when 0
    Dim TargetCreditMonth   As String
    Dim CreditMonth         As String
    
    CreditMonth = InputBox("Enter year-month for which payment taken by credit card", "Credit Period", "2023-11")
    If StrPtr(CreditMonth) = 0 Then
        Exit Sub
    End If
    
    'Save current workbook range for later when be focused on another worksheet
    Set TargetWorkbook = ActiveWorkbook
    Set TargetWorksheet = ActiveSheet
    
    'Open Bank account export data
    Workbooks.Open "d:\tmp\BankExports\Credit-card.xls"
    Set SourceDataWorkbook = ActiveWorkbook
    SourceLastRow = Cells(Rows.Count, 1).End(xlUp).row
    
    Set SourceTableRange = SourceDataWorkbook.ActiveSheet.Range(Cells(SourceFirstRowOfTable, 1), Cells(SourceLastRow, 7))
    'Set SourceTableRange = SourceDataWorkbook.ActiveSheet.Range("A1:A18")    'For debug
    
    'Go back to main workbook
    TargetWorkbook.Activate
    
    Dim SourceDate          As String
    Dim SourceOperation     As String
    Dim SourceSpent         As Variant  'Because it may be "-" when 0
    
    Dim TargetCurrentRow    As Integer
    Dim AddItemRow          As Integer
    Dim ItemAdded           As Boolean
    Dim AddItem             As Boolean
    
    TargetCurrentRow = TargetSheetStartRow
    
    '---------------------------------------------------------
    'Find row in Target workbook where insert line with data from source
    '---------------------------------------------------------
    While (CreditMonth <> TargetWorksheet.Cells(TargetCurrentRow, TargetColumnCreditMonth) And TargetCurrentRow < 1000)
        TargetCurrentRow = TargetCurrentRow + 1
    Wend
    
    '-----------------------------
    'Loop over excel exported from Bank
    '-----------------------------
    For Each rangeRow In SourceTableRange.Rows
        
        SourceDate = rangeRow.Columns(SourceColumnDate)
        SourceOperation = rangeRow.Columns(SourceColumnOperationName)
        SourceSpent = rangeRow.Columns(SourceColumnSpent)
        
        Debug.Print ("----------------- Source Item -------------------")
        Debug.Print ("SourceDate=" & SourceDate)
        Debug.Print ("SourceSpent=" & SourceSpent)
        
        'By default item will be added - prove overwise if no need
        ItemAdded = True
        AddItem = True

        'Set row for potential add item to main table
        AddItemRow = TargetCurrentRow
        
        '--------------------------------------------------------
        'Check if item already exist in Target workbook
        '--------------------------------------------------------

        Do While CreditMonth = TargetWorksheet.Cells(TargetCurrentRow, TargetColumnCreditMonth)
        
            TargetOperation = TargetWorksheet.Cells(TargetCurrentRow, TargetColumnOperationName)
            TargetSpent = TargetWorksheet.Cells(TargetCurrentRow, TargetColumnSpent)
                        
            Debug.Print ("=== Credit Month match item in Target sheet row: " & TargetCurrentRow)
            Debug.Print ("Target Date: " & TargetWorksheet.Cells(TargetCurrentRow, TargetColumnDate))
            Debug.Print ("Target Spent: " & TargetSpent)
            
            'Test if item already exist in MainTable and skip add part
            If (SourceSpent = TargetSpent And CreditMonth = TargetWorksheet.Cells(TargetCurrentRow, TargetColumnCreditMonth)) Then
                
                Debug.Print ("!!! Item already found - Set AddItem = False")
                AddItem = False
                Exit Do
                
            End If
            
            TargetCurrentRow = TargetCurrentRow + 1
            Debug.Print ("Item with same date not found - AddItem = " & AddItem)
        Loop
        
        Debug.Print ("End of Main workbook loop")
        
        '--------------------------------------------------
        'Add Item to my Excel
        '--------------------------------------------------
        If (AddItem) Then
            
            Debug.Print ("Row for insert : " & TargetCurrentRow)
            ItemAdded = True
            
            'Add new row to my excel and put all values
            TargetWorksheet.Cells(TargetCurrentRow, "A").EntireRow.Insert CopyOrigin:=xlFormatFromLeftOrAbove
            TargetWorksheet.Cells(TargetCurrentRow, TargetColumnDate).Value = SourceDate
            TargetWorksheet.Cells(TargetCurrentRow, TargetColumnOperationName).Value = SourceOperation
            TargetWorksheet.Cells(TargetCurrentRow, TargetColumnSpent).Value = SourceSpent
            TargetWorksheet.Cells(TargetCurrentRow, TargetColumnCreditMonth).Value = CreditMonth
            
            Debug.Print ("Row Add finish - Current row to check: " & TargetCurrentRow)
        End If
                
        Debug.Print ("-")
    Next rangeRow

    SourceDataWorkbook.Close SaveChanges:=False

End Sub