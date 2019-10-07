Attribute VB_Name = "boardModule"
Option Explicit

    Dim scoreBoardRange As Range
    Dim scoreboxP1 As Range, winnerRange As Range, highCardRange As Range
    Dim nameBoxer As Range, scoreBox As Range
    Dim allRanges As Range

        
Sub saveScores()

'Save each game's final scores to a table named historyTable (on sheet)


End Sub

        
Sub clearScoreboard()

    Set scoreBoardRange = ThisWorkbook.Worksheets("the House").Range("displayScoreboard")
    Set scoreboxP1 = ThisWorkbook.Worksheets("the House").Range("displayP1")
    Set nameBoxer = ThisWorkbook.Worksheets("the House").Range("firstNameBox")
   
    Set allRanges = ThisWorkbook.Worksheets("the House").Range("valueRanges")

    allRanges.ClearContents
    scoreBoardRange.ClearContents

   Range(nameBoxer, nameBoxer.Offset(0, 7)).Interior.Color = RGB(64, 64, 64) 'ClearContents

End Sub


Sub scoreBoard()
    
    Dim scoreRow As Range
    Dim y As Integer
    
    For y = 0 To playerCount
        Application.Wait (Now + TimeValue("0:00:01"))
        
        With nameBoxer.Offset(0, y - 1)
            .Interior.Color = vbRed
            .Value = resultSet(y, 1)
        End With
        
        With nameBoxer.Offset(1, y - 1)
            .Value = resulset(y, 2)
        End With
        
    Next
    
End Sub


Sub saveScoresTest()
    
    'Get table reference
    Dim r As Range
    Set r = ThisWorkbook.Worksheets("Game History").Range("historyTable")
    
    'Create new row and store it
    Dim newRow As ListRow
    Set newRow = r.ListObject.ListRows.Add
    
    'load the value from the scoreboard
    Dim points As Integer
    points = ThisWorkbook.Worksheets("the House").Range("PLACE_HOLDER").Value
            
    'Get the Column name and row number of our new row
    Dim iRow As Integer
    Dim str As String
    
    iRow = newRow.Index
    str = "Player 5"
    
    'insert value from scoreboard into the intersection of the column name and row number
    r.ListObject.ListColumns(str).DataBodyRange(iRow).Value = points


End Sub
