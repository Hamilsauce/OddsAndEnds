Option Explicit


Sub formatGridPadding()

    Range("D4:D33,E4:AG4").Select
    Range("E4").Activate
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent5
        .TintAndShade = 0.799981688894314
        .PatternTintAndShade = 0
    End With

End Sub

Sub formatCanvas()

    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlDash
        .ThemeColor = 9
        .TintAndShade = 0.399945066682943
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlDash
        .ThemeColor = 5
        .TintAndShade = 0.399945066682943
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ThemeColor = 9
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ThemeColor = 9
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With Selection.Borders(xlInsideVertical)
        .LineStyle = xlDash
        .ThemeColor = 9
        .TintAndShade = 0.399945066682943
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideHorizontal)
        .LineStyle = xlDash
        .ThemeColor = 9
        .TintAndShade = 0.399945066682943
        .Weight = xlThin
    End With
End Sub


Sub formatGridCoordinates()
    
    Dim topCoordRange As Range
    Dim leftCoordRange As Range

'Top Coord Range
    Set topCoordRange = _
        ThisWorkbook.Worksheets("Grid").Range("gridCoordTop")
        
    'background-color
    With topCoordRange.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent5
        .TintAndShade = -0.249946592608417
        .PatternTintAndShade = 0
    End With
    
    With topCoordRange
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With

    'Font
     With topCoordRange.Font
        .Name = "Segoe UI"
        .FontStyle = "Regular"
        .Size = 11
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = -0.149998474074526
    End With
    
    'Borders
    With topCoordRange.Borders
        .LineStyle = xlContinuous
        .ThemeColor = 5
        .TintAndShade = 0
        .Weight = xlThin
    End With
    
    'remove borders that don't fit

    With topCoordRange
        .Borders(xlDiagonalDown).LineStyle = xlNone
        .Borders(xlDiagonalUp).LineStyle = xlNone
        .Borders(xlEdgeLeft).LineStyle = xlNone
        .Borders(xlInsideVertical).LineStyle = xlNone
    End With

End Sub

Sub formatGridCorner()
'Need to refine/ add variables
'corner formatting
    'D3 Stuff
    Range("C3:C4,D3").Select
    Range("D3").Activate
    
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent3
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorLight1
        .TintAndShade = 0.349986266670736  'medium black
        .PatternTintAndShade = 0
    End With
    With Selection.Font
        .Name = "Segoe UI"
        .FontStyle = "Regular"
        .Size = 11
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleNone
        .ThemeColor = xlThemeColorLight1
        .TintAndShade = 0.499984741
        .ThemeFont = xlThemeFontNone
    End With
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ThemeColor = 5
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ThemeColor = 5
        .TintAndShade = 0
        .Weight = xlThin
    End With
    Selection.Borders(xlEdgeRight).LineStyle = xlNone
    Selection.Borders(xlInsideVertical).LineStyle = xlNone
    Selection.Borders(xlInsideHorizontal).LineStyle = xlNone
   
   'more
    Range("D3").Select
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    Selection.Borders(xlEdgeLeft).LineStyle = xlNone
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ThemeColor = 5
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ThemeColor = 5
        .TintAndShade = 0
        .Weight = xlThin
    End With
    Selection.Borders(xlEdgeRight).LineStyle = xlNone
    Selection.Borders(xlInsideVertical).LineStyle = xlNone
    Selection.Borders(xlInsideHorizontal).LineStyle = xlNone
  
  
  'C4 stuff
  
    Range("C4").Select
    
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ThemeColor = 5
        .TintAndShade = 0
        .Weight = xlThin
    End With

    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ThemeColor = 5
        .TintAndShade = 0
        .Weight = xlThin
    End With

End Sub
