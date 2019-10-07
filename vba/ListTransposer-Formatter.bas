Sub ListTrans()
'To do: remove the empty spaces remaining after transposing rows
    
    Dim wb As Workbook
    Dim wks As Worksheet
    
    Set wb = ThisWorkbook
    Set wks = wb.ActiveSheet
    
    Dim baseCell As Range
    Set baseCell = ActiveCell

    Dim rowNum As Integer
    rowNum = 0
    
    'optimize
        Application.ScreenUpdating = False
        Application.EnableEvents = False
        Application.DisplayAlerts = False
        Application.Calculation = xlManual
    
    'Set the environment for current pass - define current region, set up refernence cell
    On Error GoTo ErrorHandle
        
    Dim usedRange As Range
    Dim count As Long: count = 0
    Do
        Set usedRange = Range(baseCell, baseCell.End(xlDown))
      
       'Check that the macro didnt take off down the spreadhseet
        count = usedRange.Cells.count
        If count > 100000 Then
            MsgBox "shit load of cells, cancelling macro"
            Exit Sub
        Else
        
        'Fill the values of current cells with into an Array
            Dim i As Long: i = 0
            Dim r As Range
            Dim rangeVals() As String
            
            For Each r In usedRange
                ReDim Preserve rangeVals(0 To i)
                rangeVals(i) = r.Value
                i = i + 1
                
            Next r
        End If
        
        'Place those values in the row-based cells (transpose it)
        For i = LBound(rangeVals) To UBound(rangeVals)
            baseCell.Offset(0, i + 1).Value = rangeVals(i)

        Next i
        
        'Prep for next pass
        i = 0
        usedRange.ClearContents
        Erase rangeVals
        
        rowNum = rowNum + 1
        baseCell.Value = rowNum
        
        Set baseCell = baseCell.End(xlDown)

    Loop While Not baseCell.Value = ""
    
ErrorHandle:
        Application.ScreenUpdating = True
        Application.EnableEvents = True
        Application.DisplayAlerts = True
        Application.Calculation = xlAutomatic
    
End Sub