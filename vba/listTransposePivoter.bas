Sub ListTransposePivoter()
'

    Dim wb As Workbook
    Dim wks As Worksheet
    
    Set wb = ThisWorkbook
    Set wks = wb.ActiveSheet
        
    Dim r As Range

    Dim baseCell As Range
    Set baseCell = ActiveCell
    
    Debug.Print baseCell.Address
    
    
    'Loop starts here
    
    Dim rangeVals() As String
    Dim usedRange As Range
    Dim i As Long, count As Long
    
    
    Do
        Set usedRange = Range(baseCell, baseCell.End(xlDown))
        Debug.Print usedRange.Address
       
        count = usedRange.Cells.count
        If count > 1000 Then
            MsgBox "shit load of cells, cancelling macro"
            Exit Sub
        Else
        
            i = 0
            For Each r In usedRange
                
                ReDim Preserve rangeVals(0 To i)
                rangeVals(i) = r.Value
                
                i = i + 1
            Next r
            
            i = 0
            
        End If
        
        For i = LBound(rangeVals) To UBound(rangeVals)
            
            baseCell.Offset(0, i + 1).Value = rangeVals(i)
            Debug.Print rangeVals(i)
            
            i = i + 1
            
        Next i
        
        i = 0
        usedRange.ClearContents
        Erase rangeVals
        
        baseCell = baseCell.End(xlDown)
    
    Loop While Not baseCell.Value = ""
    
    
End Sub
