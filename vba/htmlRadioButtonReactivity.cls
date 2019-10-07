VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit
    Dim astrologySelection As Integer
    Dim moralReligiousSelection As Integer
    Dim charitySelection As Integer
    Dim luckSelection As Integer
    Dim religionInSchoolSelection As Integer
    Dim displayChoice As Range
    
    Function getDisplay(selection)
    Dim msg As String

    Select Case selection
        Case 1
            msg = "User Strongly Disagrees!"
        Case 2
           msg = "User Disagrees!"
        Case 3
            msg = "User Agrees!"
        Case 4
            msg = "User Strongly Agrees!"
    End Select
    
    Debug.Print msg
    getDisplay = msg
    
End Function

' astro

    Private Sub astrologyStronglydisagree_GotFocus()
        Dim choice As String
        
        astrologySelection = 1
        choice = getDisplay(astrologySelection)
        
        Set displayChoice = ActiveSheet.Range("astroCell")
        displayChoice.Select
        displayChoice.Value = choice
        
    End Sub

    Private Sub astrologyDisagree_GotFocus()
        Dim choice As String
        
        astrologySelection = 2
        choice = getDisplay(astrologySelection)
        
        Set displayChoice = ActiveSheet.Range("astroCell")
        displayChoice.Select
        displayChoice.Value = choice
        
    End Sub

    Private Sub astrologyAgree_GotFocus()
        Dim choice As String
        
        astrologySelection = 3
        choice = getDisplay(astrologySelection)
        
        Set displayChoice = ActiveSheet.Range("astroCell")
        displayChoice.Select
        displayChoice.Value = choice
        
    End Sub

    Private Sub astrologyStronglyagree_GotFocus()
        Dim choice As String
        
        astrologySelection = 4
        choice = getDisplay(astrologySelection)
        
        Set displayChoice = ActiveSheet.Range("astroCell")
        displayChoice.Select
        displayChoice.Value = choice

    End Sub


' moralReligious

    Private Sub moralReligiousStronglydisagree_GotFocus()
        Dim choice As String
        
        moralReligiousSelection = 1
        choice = getDisplay(moralReligiousSelection)
        
        Set displayChoice = ActiveSheet.Range("moralCell")
        displayChoice.Select
        displayChoice.Value = choice
        
    End Sub

    Private Sub moralReligiousDisagree_GotFocus()
        Dim choice As String
        
        moralReligiousSelection = 2
        choice = getDisplay(moralReligiousSelection)
        
        Set displayChoice = ActiveSheet.Range("moralCell")
        displayChoice.Select
        displayChoice.Value = choice
        
    End Sub

    Private Sub moralReligiousAgree_GotFocus()
        Dim choice As String
        
        moralReligiousSelection = 3
        choice = getDisplay(moralReligiousSelection)
        
        Set displayChoice = ActiveSheet.Range("moralCell")
        displayChoice.Select
        displayChoice.Value = choice
        
    End Sub

    Private Sub moralReligiousStronglyagree_GotFocus()
        Dim choice As String
        
        moralReligiousSelection = 4
        choice = getDisplay(moralReligiousSelection)
        
        Set displayChoice = ActiveSheet.Range("moralCell")
        displayChoice.Select
        displayChoice.Value = choice

    End Sub


' charity

    Private Sub charityStronglydisagree_GotFocus()
        Dim choice As String
        
        charitySelection = 1
        choice = getDisplay(charitySelection)
        
        Set displayChoice = ActiveSheet.Range("charityCell")
        displayChoice.Select
        displayChoice.Value = choice
        
    End Sub

    Private Sub charityDisagree_GotFocus()
        Dim choice As String
        
        charitySelection = 2
        choice = getDisplay(charitySelection)
        
        Set displayChoice = ActiveSheet.Range("charityCell")
        displayChoice.Select
        displayChoice.Value = choice
        
    End Sub

    Private Sub charityAgree_GotFocus()
        Dim choice As String
        
        charitySelection = 3
        choice = getDisplay(charitySelection)
        
        Set displayChoice = ActiveSheet.Range("charityCell")
        displayChoice.Select
        displayChoice.Value = choice
        
    End Sub

    Private Sub charityStronglyagree_GotFocus()
        Dim choice As String
        
        charitySelection = 4
        choice = getDisplay(charitySelection)
        
        Set displayChoice = ActiveSheet.Range("charityCell")
        displayChoice.Select
        displayChoice.Value = choice

    End Sub


' luck

    Private Sub luckStronglydisagree_GotFocus()
        Dim choice As String
        
        luckSelection = 1
        choice = getDisplay(luckSelection)
        
        Set displayChoice = ActiveSheet.Range("luckCell")
        displayChoice.Select
        displayChoice.Value = choice
        
    End Sub

    Private Sub luckDisagree_GotFocus()
        Dim choice As String
        
        luckSelection = 2
        choice = getDisplay(luckSelection)
        
        Set displayChoice = ActiveSheet.Range("luckCell")
        displayChoice.Select
        displayChoice.Value = choice
        
    End Sub

    Private Sub luckAgree_GotFocus()
        Dim choice As String
        
        luckSelection = 3
        choice = getDisplay(luckSelection)
        
        Set displayChoice = ActiveSheet.Range("luckCell")
        displayChoice.Select
        displayChoice.Value = choice
        
    End Sub

    Private Sub luckStronglyagree_GotFocus()
        Dim choice As String
        
        luckSelection = 4
        choice = getDisplay(luckSelection)
        
        Set displayChoice = ActiveSheet.Range("luckCell")
        displayChoice.Select
        displayChoice.Value = choice

    End Sub


'religionInSchool

    Private Sub religionInSchoolStronglydisagree_GotFocus()
        Dim choice As String
        
        religionInSchoolSelection = 1
        choice = getDisplay(religionInSchoolSelection)
        
        Set displayChoice = ActiveSheet.Range("religionInSchoolCell")
        displayChoice.Select
        displayChoice.Value = choice
        
    End Sub

    Private Sub religionInSchoolDisagree_GotFocus()
        Dim choice As String
        
        religionInSchoolSelection = 2
        choice = getDisplay(religionInSchoolSelection)
        
        Set displayChoice = ActiveSheet.Range("religionInSchoolCell")
        displayChoice.Select
        displayChoice.Value = choice
        
    End Sub

    Private Sub religionInSchoolAgree_GotFocus()
        Dim choice As String
        
        religionInSchoolSelection = 3
        choice = getDisplay(religionInSchoolSelection)
        
        Set displayChoice = ActiveSheet.Range("religionInSchoolCell")
        displayChoice.Select
        displayChoice.Value = choice
        
    End Sub

    Private Sub religionInSchoolStronglyagree_GotFocus()
        Dim choice As String
        
        religionInSchoolSelection = 4
        choice = getDisplay(religionInSchoolSelection)
        
        Set displayChoice = ActiveSheet.Range("religionInSchoolCell")
        displayChoice.Select
        displayChoice.Value = choice

    End Sub



