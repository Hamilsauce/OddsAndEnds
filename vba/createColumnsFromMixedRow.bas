Option Explicit

    Dim tableRange As Range
    Dim currentCell As Range
    Dim targetCell As Range
    Dim cell As Range
    Dim wks As Worksheet

Sub gameTime()

    Dim tableRange As Range
    Dim currentCell As Range
    Dim targetCell As Range
    Dim cell As Range
    Dim wks As Worksheet

    Set wks = ThisWorkbook.Worksheets("Sheet2")
    Set tableRange = wks.UsedRange

    For Each cell In tableRange
        If Not cell Is Nothing Then
            Set currentCell = cell
            Debug.Print cell.Address
            cell.Offset(0, 1).Value = cell.Offset(1, 0).Value
            cell.Offset(1, 0).Delete Shift:=xlUp
        End If
    Next cell

    cell = Nothing
    Call Module1.formatAsTable
End Sub

Sub formatAsTable()

    Dim gameTable As ListObject
    Dim headerRange As Range
    Dim counter As Integer: counter = 0

    Set gameTable = wks.ListObjects.Add(xlSrcRange, tableRange, , xlNo).Name = "fuck"
    Set headerRange = gameTable.HeaderRowRange

    For Each cell In headerRange
        If counter = 0 Then
            cell.Value = "Game"
            counter = counter + 1
        ElseIf counter > 0 Then
            cell.Value = "Hours Played"
        End If
    Next cell

    cell = Nothing

    wks.Range("A1").Activate

End Sub
