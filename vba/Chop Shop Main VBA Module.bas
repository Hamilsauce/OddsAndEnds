Attribute VB_Name = "ColumnMoverCode"
Option Explicit
Sub ReOrder()

    Dim wbWorkFile As Workbook
    Dim wbRoster As Workbook
    Dim New_wbCompare As Workbook
        
    Dim wkSheets As Worksheet
    Dim wksInputSheet As Worksheet
    Dim wksOutputSheet As Worksheet
    Dim wksControls As Worksheet
    Dim wksData As Worksheet
    Dim New_Compare
        
    Dim loNewRoster As ListObject
        
    Dim rngUnOrderedHeaders As Range
    Dim rngCorrectHeaders As Range
    Dim hCell As Range
    Dim rngFound As Range
        
    Dim k As Long: k = 0
    Dim iCol As Long
    Dim hCounter As Long
    Dim wkCounter As Long

    Dim arrHeaderList() As String
    Dim strFound As String
    
'The Main Attraction - Also checks for data and worksheets. Formats ReOrder data range as table

    Set wbWorkFile = ThisWorkbook
    Set wksControls = wbWorkFile.Worksheets("Controls")
    Set wksInputSheet = wbWorkFile.Worksheets("Input Sheet")
    Set wksData = wbWorkFile.Worksheets("Data")
    
    Application.ScreenUpdating = False

'***Assign ReOrdered Variable - Check if "ReOrdered" Sheet Exists, create If Not; If So, Clear sheet cells of content.

    On Error Resume Next
    If ThisWorkbook.Worksheets("Output Sheet") Is Nothing Then
        Set wksOutputSheet = wbWorkFile.Worksheets.Add(, After:=wksControls)
        wksOutputSheet.Name = "Output Sheet"
        
        On Error GoTo 0
    Else
        Set wksOutputSheet = wbWorkFile.Worksheets("Output Sheet")
        wksOutputSheet.Cells.Delete
    End If
    
'***Make Sure Roster is in Roster Sheet

    If wksInputSheet.Range("A1").Value = "" Then
        MsgBox "Roster needs to be in Roster Input Sheet starting at cell A1 and no Empty Cells in Headers. Try again."
        wksControls.Activate
        Exit Sub
    Else
        wksInputSheet.Activate
        Set rngUnOrderedHeaders = wksInputSheet.Range("A1", Range("A1").End(xlToRight))
    End If
    
'***Get Header Values into Array

    Set rngCorrectHeaders = Range(Range("nrPivotStart") _
        .Offset(1, 0), Range("nrPivotStart").Offset(1, 0).End(xlDown)) '.SpecialCells(xlCellTypeVisible)
    
    For Each hCell In rngCorrectHeaders
    
        If Not hCell Is Nothing Then
            ReDim Preserve arrHeaderList(k)
            arrHeaderList(k) = hCell.Value
            k = k + 1
        Else
            MsgBox "There is a blank cell in the headers range. Please Correct."
            Exit Sub
        End If
        
    Next hCell

'***Prep Roster - Remove Spaces in Roster Headers, Remove Extra Junk in Headers, rename some headers
    
   On Error Resume Next

   rngUnOrderedHeaders.Replace What:=" ", Replacement:="", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    
    'Check for Double spaces
    rngUnOrderedHeaders.Replace What:="  ", Replacement:="", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False

    rngUnOrderedHeaders.Replace What:="IndividualNPI", Replacement:="PractitionerNPI", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    
    rngUnOrderedHeaders.Replace What:="NPI", Replacement:="PractitionerNPI", LookAt:=xlWhole, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    
    rngUnOrderedHeaders.Replace What:="*Group*NPI*", Replacement:="ProviderNPI", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    
    rngUnOrderedHeaders.Replace What:="*GNPI*", Replacement:="ProviderNPI", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
        
      rngUnOrderedHeaders.Replace What:="*Tax*", Replacement:="TIN", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
        
    rngUnOrderedHeaders.Replace What:="Address*1", Replacement:="LocationAddressLine1", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False

    rngUnOrderedHeaders.Replace What:="Address*2", Replacement:="LocationAddressLine2", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    
    rngUnOrderedHeaders.Replace What:="remittance", Replacement:="Billing", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    
    rngUnOrderedHeaders.Replace What:="Billing*1", Replacement:="BillingAddressLine1", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
        
    rngUnOrderedHeaders.Replace What:="Billing*2", Replacement:="BillingAddressLine2", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
        
    rngUnOrderedHeaders.Replace What:="BillingLocation", Replacement:="Billing", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
        
    rngUnOrderedHeaders.Replace What:="*Specialist*", Replacement:="Hatcode", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    
    rngUnOrderedHeaders.Replace What:="*State", Replacement:="LocationState", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    
    
    On Error GoTo 0

'***Begin Search and Copy

    For hCounter = LBound(arrHeaderList) To UBound(arrHeaderList)
    
        On Error Resume Next
        
        iCol = hCounter
        
        Set rngFound = rngUnOrderedHeaders _
            .Find(What:=arrHeaderList(hCounter), LookIn:=xlValues, _
             LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
             MatchCase:=False, SearchFormat:=False)
            
       On Error GoTo 0
       
        If rngFound Is Nothing Then
            wksOutputSheet.Range("A1").Offset(, iCol).Value = "$<>" & arrHeaderList(hCounter)
        Else
            rngFound.EntireColumn.Copy
            
            wksOutputSheet.Activate
            Range("A1").Offset(, iCol).Select
            wksOutputSheet.Paste
            rngFound.Interior.Color = rgbOrange
            
        End If
        
    Next hCounter
   
'***Create Table out of new columns, name data range as ComparisonData1
    wksOutputSheet.Activate
    wksOutputSheet.UsedRange.Select
    wksOutputSheet.ListObjects.Add( _
       xlSrcRange, wksOutputSheet.UsedRange, , xlYes).Name = "ReOrderTable"
    Set loNewRoster = wksOutputSheet.ListObjects("ReOrderTable")
    loNewRoster.TableStyle = "TableStyleMedium18"

'***Call LocationKey Code

    
    Application.ScreenUpdating = True


End Sub


Sub RemoveNonMatchColumns()

    Dim wbWorkFile As Workbook
    Dim wkSheets As Worksheet
    Dim wksInputSheetSheet As Worksheet
    Dim wksOutputSheet As Worksheet
    Dim wksControls As Worksheet
    Dim wksData As Worksheet
    Dim loNewRoster As ListObject
    Dim hCell As Range
    Dim rngFound As Range
    Dim strFound As String
    Dim rngReOrderHeaders As Range
    
    Set wbWorkFile = ThisWorkbook
    Set wksControls = wbWorkFile.Worksheets("Controls")
    Set wksInputSheetSheet = wbWorkFile.Worksheets("Input Sheet")
    Set wksOutputSheet = wbWorkFile.Worksheets("Output Sheet")
    
    Application.ScreenUpdating = False
        
    wksOutputSheet.Activate
    Set rngReOrderHeaders = wksOutputSheet.Range("A1", Range("A1").End(xlToRight))
               
    If rngReOrderHeaders.Text = "" Then
        MsgBox "No data in sheet to Trim. Maybe you were looking for the button just above?"
        Exit Sub
    Else
    
    End If

    For Each hCell In rngReOrderHeaders

        If IsEmpty(hCell) = True Then
            hCell.EntireColumn.Delete Shift:=xlToLeft
        Else
            On Error Resume Next
            Set rngFound = rngReOrderHeaders _
                .Find(What:="$<>", LookIn:=xlValues, _
                    LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
                    MatchCase:=False, SearchFormat:=False)
            
            rngFound.EntireColumn.Delete Shift:=xlToLeft
        End If
        
    Next hCell

    Application.ScreenUpdating = True

End Sub



Sub CFormatCompare()

        Range("B2").Select
    Selection.End(xlUp).Select
    Selection.End(xlToLeft).Select
    Range("ReOrderTable[#All]").Select
    Sheets("Input Sheet").Select
    
    Range("A1").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    ActiveWorkbook.Names.Add Name:="CompareRange", RefersToR1C1:= _
        "=Roster_Input!R1C1:R2C132"
    ActiveWorkbook.Names("CompareRange").Comment = ""
    Sheets("Output Sheet").Select
    Range("ReOrderTable[#All]").Select
    Cells.FormatConditions.Delete
    Range("ReOrderTable[#All]").Select
    Selection.FormatConditions.Add Type:=xlExpression, Formula1:= _
        "=COUNTIF(CompareRange,A1)=0"
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .Color = 65535
        .TintAndShade = 0
    End With
    Selection.FormatConditions(1).StopIfTrue = False
    
End Sub


Sub ClearRoster()

    ThisWorkbook.Worksheets("Input Sheet").Cells.Delete
    
    If Worksheets("Input Sheet").Cells.Text = "" Then
        MsgBox "Consider the Roster Cleared."
    Else
    
    End If

End Sub


