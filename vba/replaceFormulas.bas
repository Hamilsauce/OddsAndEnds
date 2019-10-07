Option Explicit

Sub toNums()

Dim wb As Workbook
Dim sheets As Worksheets
Dim wks As Worksheet
Dim r As Range

    Set wb = ThisWorkbook
    
   ' Set sheets = wb.Worksheets

    For Each wks In wb.Worksheets
        Set r = wks.UsedRange
        Debug.Print r.Address

        With r
            .copy
            .PasteSpecial Paste:=xlPasteValues
        End With
    
    Next wks

    MsgBox "sub is done running, go see how it went."
        

End Sub

