Option Explicit

Dim wb As Workbook
Dim graphSheet As Worksheet

Dim zoomLevels(1) As Integer
Dim initTest As Boolean


Sub getVariables()
'run this immediately at workbook.Open
    
    Set wb = ThisWorkbook
    Set graphSheet = wb.Worksheets("Close Price Graph")
   
    'set zooms in array
    Dim i As Variant, k As Integer: k = 70
    For i = 0 To 1
        zoomLevels(i) = k
        k = k + 20

    Next i
    
    initTest = True
    
End Sub

Sub goToFullView(holler)

    'test if variables are set, if not run sub
    If initTest = False Then
        Call getVariables
    Else

    End If
    
    On Error GoTo errorHandle
       
    Dim fullView As Range
    Dim fullViewName As String
    
    Set fullView = graphSheet.Range("fullViewRange")
    fullViewName = fullView.Name
    
    Application.ScreenUpdating = False
    
    'go to view range and zoom apprpriately
    Application.Goto reference:=fullViewName
    ActiveWindow.Zoom = zoomLevels(1)
    
    Application.ScreenUpdating = True
    Exit Sub
    
errorHandle:
    MsgBox "that didn't work. probably no named range for fullView"
    Application.ScreenUpdating = True
    
End Sub


Sub goToGraphView()

    'test if variables are set, if not run sub
    If initTest = False Then
        Call getVariables
    Else
    End If
    
    Dim graphView As Range
    Dim graphViewName As String
        
    Set graphView = graphSheet.Range("graphRange")
    Debug.Print graphViewName = graphView.Name
    
    Application.ScreenUpdating = False

    'go to view range and zoom apprpriately
    Application.Goto reference:=graphViewName
    ActiveWindow.Zoom = zoomLevels(0)
    
    Application.ScreenUpdating = True
    Exit Sub
    
errorHandle:
    MsgBox "that didn't work. probably no named range for graphView"
    Application.ScreenUpdating = True

End Sub
