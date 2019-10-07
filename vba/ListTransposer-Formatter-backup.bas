

Sub ListTrans()
'To do: remove the empty spaces remaining after transposing rows
    
    Dim wb As Workbook
    Dim wks As Worksheet
    
    Set wb = ThisWorkbook
    Set wks = wb.ActiveSheet
        

    Dim baseCell As Range
    Set baseCell = ActiveCell
    
    Debug.Print baseCell.Address
    
    'Loop starts here
    Dim rangeVals() As String
    Dim usedRange As Range
    Dim i As Long, count As Long
    
    Dim rowNum As Integer
    
    'Set the environment for current pass - define current region, set up refernence cell
    Do
        Set usedRange = Range(baseCell, baseCell.End(xlDown))
        usedRange.Select
        Debug.Print usedRange.Address
       
       'Check that the macro didnt take off down the spreadhseet
        count = usedRange.Cells.count
        If count > 100000 Then
            MsgBox "shit load of cells, cancelling macro"
            Exit Sub
        Else
        
        'Fill the values of current cells with into an Array
            i = 0
            
            Dim r As Range
            For Each r In usedRange
                ReDim Preserve rangeVals(0 To i)
               ' rangeVals(i).Select
                rangeVals(i) = r.Value
                i = i + 1
            Next r
        End If
        
        'Place those values in the row-based cells (transpose it)
    
    
        For i = LBound(rangeVals) To UBound(rangeVals)
            baseCell.Offset(0, i + 1).Select
            baseCell.Offset(0, i + 1).Value = rangeVals(i)
            
           ' Debug.Print rangeVals(i)
           ' i = i + 1
        Next i
        i = 0
        
        usedRange.ClearContents
        Erase rangeVals
        
        rowNum = rowNum + 1
        baseCell.Value = rowNum
        
        baseCell.End(xlDown).Select
        Set baseCell = baseCell.End(xlDown)
        
        Debug.Print baseCell.Address
        
    Loop While Not baseCell.Value = ""
    
    
End Sub