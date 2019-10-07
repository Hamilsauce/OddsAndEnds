Attribute VB_Name = "functions"
Function getCardValue(playerCount) As Integer

    Randomize
    getCardValue = Int((playerCount - 1 + 1) * Rnd + 1)
        
End Function

Function inputTest()
    Dim Temp$
     
    Temp = InputBox("Enter something here:", "Inputbox")
    If StrPtr(Temp) = 0 Then
        MsgBox "You pressed Cancel!" 'Option 1
    Else
        If Temp = "" Then
            MsgBox "You entered nothing and pressed OK" ' Option 2
        Else
            MsgBox Temp 'Option 3
        End If
    End If
End Function

