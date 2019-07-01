Attribute VB_Name = "MacroCode"
Option Explicit

Sub MacroListForm()
Attribute MacroListForm.VB_ProcData.VB_Invoke_Func = "Q\n14"
   ' If toggleMacroForm = False Then
        MacroShortcuts.Show

End Sub
Sub listBoxFixer()

    Dim wb As Workbook
    Dim wks As Worksheet
    Dim macroList As Range
    Dim count As Integer
    Dim rSource As String
    
    Set wb = ThisWorkbook
    Set wks = wb.Worksheets("MacroList")
    Set macroList = wks.Range("MacroTable").ListObject.DataBodyRange
    
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.DisplayAlerts = False
    Application.Calculation = xlManual
    
    MacroShortcuts.Enabled = True
    MacroShortcuts.ListBox2.Enabled = True

    On Error Resume Next
    
    rSource = "TableRange"
    MacroShortcuts.ListBox2.RowSource = rSource
    
    count = MacroShortcuts.ListBox2.ListCount
    Debug.Print (count)

    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Application.DisplayAlerts = True
    Application.Calculation = xlManual

End Sub
Sub ToggleUI()
Attribute ToggleUI.VB_ProcData.VB_Invoke_Func = "W\n14"
'Ctrl Shft W
'Switches between 3 window states

    'Test if any of the UI is switched off and store result in Toggle
    Dim bToggle As Boolean
    Dim winState As Integer
    
    If Application.DisplayFullScreen = True Then
            winState = 3
        
    ElseIf Application.DisplayFormulaBar = False _
        Or ActiveWindow.DisplayHorizontalScrollBar = False _
        Or ActiveWindow.DisplayVerticalScrollBar = False _
        Or ActiveWindow.DisplayWorkbookTabs = False _
        And Application.DisplayFullScreen = False _
    Then
            winState = 2
        
    ElseIf Application.DisplayFormulaBar = True _
        Or ActiveWindow.DisplayHorizontalScrollBar = True _
        Or ActiveWindow.DisplayVerticalScrollBar = True _
        Or ActiveWindow.DisplayWorkbookTabs = True _
        And Application.DisplayFullScreen = False _
    Then
            winState = 1
        
    End If
    
    Application.ScreenUpdating = False
    
            'Given current state determined by above tests, move to the preceding window state in
            'toggle order (1 = normal; 2 = collapsed Ribbon/formula bar, headers, etc; 3 = no UI
    Select Case winState
        
        Case 1
            CommandBars.ExecuteMso "MinimizeRibbon"
            
            Application.DisplayFullScreen = False
            Application.DisplayFormulaBar = False
            
            With ActiveWindow
                '.DisplayHorizontalScrollBar = False
                '.DisplayVerticalScrollBar = False
                .DisplayHeadings = False
                .DisplayWorkbookTabs = False
            End With
            
         Case 2
                    'check current window state (normal or fullscreen/max) and if normal,unmaximize after fullscreen is triggered
            If ActiveWindow.WindowState = xlNormal Then
                Application.DisplayFullScreen = True
                ActiveWindow.WindowState = xlNormal
                
            Else
                Application.DisplayFullScreen = True
                
            End If
        
        Case 3
            Application.DisplayFullScreen = False
            
            CommandBars.ExecuteMso "MinimizeRibbon"
            Application.DisplayFormulaBar = True
            
            With ActiveWindow
                .DisplayHorizontalScrollBar = True
                .DisplayVerticalScrollBar = True
                .DisplayHeadings = True
                .DisplayWorkbookTabs = True
            End With
            
        Case Else
            MsgBox "winstate = " & winState
            
        End Select
        
    Application.ScreenUpdating = True

    
End Sub

Sub createNewWB()
Attribute createNewWB.VB_ProcData.VB_Invoke_Func = "N\n14"
'Ctrl Shft N
'Creates a new WB like Ctrl N, but with preferred custom settings

    Dim newWB As Workbook
    
    Application.DisplayAlerts = False
    Application.ScreenUpdating = False
    
    Set newWB = Workbooks.Add
    
    With Application
        .WindowState = xlNormal
        .FormulaBarHeight = 1
        .AutoRecover.Time = 2
    End With
    
    With ActiveWindow
        .DisplayGridlines = False
        .DisplayHorizontalScrollBar = True
        .DisplayVerticalScrollBar = True
        .DisplayHeadings = True
        .DisplayWorkbookTabs = True
    End With
    
    CommandBars.ExecuteMso "MinimizeRibbon"
    ActiveSheet.DisplayPageBreaks = False
    
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    
End Sub

Sub OpenLoginFile()
Attribute OpenLoginFile.VB_ProcData.VB_Invoke_Func = "L\n14"

    Dim fso As Object
    Dim strPersonalFolder As String
    Dim strLoginFilePath As String
    Dim rngCloudCheck As Range
    
    Set fso = CreateObject("Scripting.FileSystemObject")
   ' Set rngCloudCheck = ThisWorkbook.Worksheets("Macrolist").Range("nrCloudCheck")
    strPersonalFolder = fso.GetParentFolderName(ThisWorkbook.Path)
    
    On Error GoTo Error
    
    Workbooks.Open (strPersonalFolder & "\" & "LoginsPersonal.xlsm")
    
    Exit Sub
    
Error:
    MsgBox "Excel Shortcuts file not found at " & _
        strPersonalFolder & "\" & "LoginsPersonal.xlsm." & _
        " Check the Autoload Folder."

End Sub

Sub OpenExcelShortcutsFile()
Attribute OpenExcelShortcutsFile.VB_ProcData.VB_Invoke_Func = "K\n14"

    Dim fso As Object
    Dim strPersonalFolder As String
    Dim strLoginFilePath As String
    Dim rngCloudCheck As Range
    
    Set fso = CreateObject("Scripting.FileSystemObject")
   ' Set rngCloudCheck = ThisWorkbook.Worksheets("Macrolist").Range("nrCloudCheck")
    strPersonalFolder = fso.GetParentFolderName(ThisWorkbook.Path)
    
    On Error GoTo Error
    
    Workbooks.Open (strPersonalFolder & "\" & "ExcelShortcutsPersonal.xlsm")
    
    Exit Sub
    
Error:
    MsgBox "Excel Shortcuts file not found at " & _
        strPersonalFolder & "\" & "ExcelShortcutsPersonal.xlsm." & _
        " Check the Autoload Folder."


End Sub

Sub HiLightDupes()
Attribute HiLightDupes.VB_ProcData.VB_Invoke_Func = "I\n14"
'Filter for Records with Empty Contract Name + TIN, delete result, remove filter and sort
'Ctrl Shift I

    Dim InitialSelect As Range
    
    Set InitialSelect = Selection
    
   ' Cells.FormatConditions.Delete

    
    InitialSelect.FormatConditions.AddUniqueValues
    InitialSelect.FormatConditions(Selection.FormatConditions.count).SetFirstPriority
    InitialSelect.FormatConditions(1).DupeUnique = xlDuplicate
    
    With InitialSelect.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent4
        .TintAndShade = 0.399945066682943
    End With
    
    InitialSelect.FormatConditions(1).StopIfTrue = False
    
End Sub


Sub NumberFormatFix()
Attribute NumberFormatFix.VB_ProcData.VB_Invoke_Func = "T\n14"
'Ctrl Shift T

    Dim TargetCol As Range
    
    Set TargetCol = Selection
    
    On Error Resume Next
    
    With TargetCol
        .NumberFormat = "0"
        .TextToColumns Destination:=TargetCol
    End With

End Sub

Sub PageBreaks()
Attribute PageBreaks.VB_ProcData.VB_Invoke_Func = "P\n14"

    Dim wks As Worksheet
    Dim sheets As Worksheets
    Dim sheetCount As Integer
    Dim i As Integer
    Dim wb As Workbook
    
    Set wb = ActiveWorkbook

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.DisplayAlerts = False
    
    
    sheetCount = wb.Worksheets.count
    
    For Each wks In Worksheets
    
        If Not wks.Visible = xlSheetHidden _
            And Not wks.Visible = xlSheetVeryHidden _
        Then
            wks.DisplayPageBreaks = False
            i = i + 1
        End If
        
    Next wks
    
    If i = sheetCount Then
        wb.Save
        MsgBox i & " = " & sheetCount
    End If
    

    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Application.DisplayAlerts = True
    

End Sub

Sub PasteValues()
Attribute PasteValues.VB_ProcData.VB_Invoke_Func = "V\n14"
'Ctrl+Shift+V
'Pastes clipboard as Values or as unformatted if pasting from outside of Excel
    
    On Error Resume Next
    
    Selection.PasteSpecial Paste:=xlPasteValues
    ActiveSheet.PasteSpecial Format:="Text", Link:=False, DisplayAsIcon:=False

End Sub

Sub PasteFormat()
Attribute PasteFormat.VB_ProcData.VB_Invoke_Func = "F\n14"
'Ctrl+Shift+F

    On Error Resume Next
    Selection.PasteSpecial Paste:=xlPasteFormats

End Sub

Sub CenterAcrossCells()
Attribute CenterAcrossCells.VB_ProcData.VB_Invoke_Func = "C\n14"
'Shortcut: Ctrl + Shift + C

    With Selection
        .HorizontalAlignment = xlCenterAcrossSelection
        .MergeCells = False
    End With
    
End Sub

Sub ClearFormats()
Attribute ClearFormats.VB_ProcData.VB_Invoke_Func = "D\n14"
'Shortcut: Ctrl+Shift+D
    
    Selection.ClearFormats

End Sub



Sub AutofitRowsCols()
Attribute AutofitRowsCols.VB_ProcData.VB_Invoke_Func = "O\n14"
'Ctrl+Shift O
'Autofits row and column height/width of each row/column in used range of sheet.
    Dim r As Range
    Dim cell As Range
    Dim i As Integer: i = 0
    
    Set r = ActiveSheet.UsedRange.SpecialCells(xlCellTypeVisible)
    
    For Each cell In r
        If Not i > 5000 Then
            cell.EntireColumn.WrapText = False
            cell.EntireColumn.AutoFit
            cell.EntireRow.AutoFit
            i = i + 1
        ElseIf i > 5000 Then
            MsgBox "5000"
            Exit For
        End If
        
    Next cell
    
End Sub

Sub zoom100()
'Decommissioned

'Zoom view back to 100%

    If (ActiveWindow.Zoom < 100) Then
        ActiveWindow.Zoom = 100
    ElseIf (ActiveWindow.Zoom > 100) Then
        ActiveWindow.Zoom = 100
    End If

End Sub

Sub Insert_RowColInserter_Showform()

'Launches userform for user input on defining number of rows or cols to insert
    
    frm_RowColInsert.Show
End Sub



