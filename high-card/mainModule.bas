Attribute VB_Name = "mainModule"
Option Explicit

'Dim playerIDs() As String 'Named by order of play
'Dim cardValues() As Integer

Dim resultSet() As Variant ' playerID with Values
Dim scoreboxP1 As Range
Dim previousCardValue As Integer
Dim playerCount As Integer 'Counter for the loop through array
Dim winnerName As String

 Dim nameBoxRange As Range, scoreBox As Range


Sub refreshValues()
    
    playerCount = 0
    previousCardValue = 0
    winnerName = ""
    Erase resultSet()

End Sub
Sub saveScores()
'Notes: Need to incorporate array
'Need to move this to new module or end of this one
    
    
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

Sub getPlayerCount()

    Dim errorCheck As Boolean
    
    Call refreshValues
    Call boardModule.clearScoreboard
    
    Do
        errorCheck = False  'Error handler checks for valid user input. if an error is encountered during the loop, the handler sets error checker to True,
                            'which then triggers the loop while condition. Each time the loop starts it resets checker to 0.
        On Error GoTo inputError
        Dim Temp
        Temp = Application.InputBox( _
            "How many players will be playing?", "Players", "Number of players...")
        
        Select Case Temp
        
            Case False      'End Game if player hits cancel button
                MsgBox ("Game Cancelled. Come back soon now!")
                Exit Sub
                
            Case ""     'Handle if player enters nothing and hits OK
                MsgBox ("Enter a number or press cancel.")
                errorCheck = True
                
            Case Is > 8     'handle if player enters count exceeding 8 limit
                If Temp > 8 Then
                    MsgBox "You entered " & Temp & " but the player limit is 8. " _
                        & "8 players will be dealt a card."
                    playerCount = 8
                End If
                
            Case Is <= 1
                MsgBox "Gotta have more than one or no players dumbass!"
                
                errorCheck = True
            Case Else
                playerCount = Temp
                
                Dim displayCount As Range
                Set displayCount = ThisWorkbook.Worksheets("the House") _
                    .Range("countDisplayRange")
                    
                displayCount.Value = playerCount
                
            End Select
            
    Loop While errorCheck = True
    
    On Error GoTo 0

    GoTo Done
    
inputError:

    MsgBox ("Please enter a number only! Must be 8 or less.")
    errorCheck = True
    
    Resume Next
    
Done:
    Call getArray

End Sub

Sub getArray()

    Dim pCounter As Integer: pCounter = 0
    Dim i As Long
    Dim i2 As Long
    Dim k As Integer
    
    If playerCount > 1 Then
        ReDim resultSet(1 To playerCount, 1 To 2)
    Else
        'Goto Code to handle only 1 player
    End If
    
   ' Call addCountToArray

    'Essentially, i = number of players playing (that account for first array dimension)
    'and i2 = the number of card values per player (always 1).
    'So we always loop down dimension 1 according to # of players, but always only loop out to the 2nd dimension 1 time

    Set scoreboxP1 = ThisWorkbook.Worksheets("the House").Range("displayP1")
    previousCardValue = 0
    i2 = 2
    
    For i = LBound(resultSet) To UBound(resultSet)
        resultSet(i, 1) = "Player " & i '+ 1
        
        Do  'Loop getting the cardValue until the returned value is not same as previous card's value
            resultSet(i, 2) = getCardValue(playerCount)
                
        Loop While resultSet(i, 2) - previousCardValue = 0
            
        Application.Wait (Now + TimeValue("0:00:01"))
        
        previousCardValue = resultSet(i, 2)
        pCounter = pCounter + 1
        
        Debug.Print resultSet(i, 1)
        Debug.Print resultSet(i, 2)
        
    Next i
    
    Call scoreBoard


End Sub


Sub scoreBoard()
    
    Dim scoreRow As Range
    Dim y As Integer: y = 0

    Set nameBoxRange = ThisWorkbook.Worksheets("the House").Range("firstNameBox")
     
     For y = 1 To playerCount
    Application.Wait (Now + TimeValue("0:00:01"))
   ' nameBoxRange.Offset(0, 0).Interior.Color = RGB(64, 64, 64)
        With nameBoxRange.Offset(0, y - 1)
            .Interior.Color = vbRed
            .Value = resultSet(y, 1)
        End With
        
        With nameBoxRange.Offset(1, y - 1)
            .Value = resultSet(y, 2)
        End With
        
    Next
    'Set scoreRow = sht.Cells(7, sht.Columns.Count).End(xlToLeft).Column
     

'
'     Set scoreboxP1 = ThisWorkbook.Worksheets("the House").Range("displayP1")
'     Set scoreRow = ThisWorkbook.Worksheets("the House").Range(Range("displayP1"), Range("displayP1").Offset(0, playerCount + 1))
'     With scoreRow
'        .ClearContents
'        .Offset(-1, 0).ClearContents
'    End With

    
'    Application.Wait (Now + TimeValue("0:00:01"))
'    With scoreboxP1.Offset(-1, 0)
'        .Value = "Player 1"
'        .Interior.Color = vbRed
'    End With
    
    Call getWinner

     
End Sub

Sub generateCardValues()
    
    resultSet(0, 0) = getCardValue(playerCount)
    Debug.Print resultSet(0, 0)




End Sub


Sub pairValuesToNames()

End Sub


Sub getWinner()
    'Match highest card to player (by way of index - P1 gets value of first generated card value),
    'Display the result
    Dim tieCount As Integer
    Dim highScore As Integer
    Dim i As Variant
    Dim winnerNames() As Variant
    
    Debug.Print "Tie Check"
    
    WorksheetFunction.Max (resultSet)
    
    For i = LBound(resultSet) To UBound(resultSet)
        If resultSet(i, 2) = WorksheetFunction.Max(resultSet) Then
    
            ReDim Preserve winnerNames(0 To tieCount)
            winnerNames(tieCount) = resultSet(i, 1)
            tieCount = tieCount + 1
        Else
        
        End If
    
    Next i

    Dim winners As String
    
    winners = Join(winnerNames, ", ")
    
    ThisWorkbook.Worksheets("the House").Range("winnerDisplay").Value = winners
    
    If UBound(winnerNames) - LBound(winnerNames) + 1 > 1 Then
        MsgBox "Tie between " & winners & "!"
    Else
        MsgBox "The Winner is " & winners & "!"
    End If

  '  Call saveScores


End Sub

Sub cleanup()

    'need to clear all players from board and all scores
    'need to make sure all values from previous game are cleaned from memory
    'Reset Player Count display

End Sub
