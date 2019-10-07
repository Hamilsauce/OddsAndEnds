Attribute VB_Name = "fileMoverCodeVBA"
Function file_exists(fl_path As String) As String
    If Dir(fl_path) <> "" And fl_path <> "" Then
        file_exists = "Exists"
    Else
        file_exists = "Not Exists"
    End If
End Function


Function folder_exists(fld_path As String) As String

    If Len(Dir(fld_path, vbDirectory)) <> 0 And fld_path <> "" Then
        folder_exists = "Exists"
    Else
        folder_exists = "Not Exists"
    End If
    
End Function


Sub move_file()

    Dim filenm As String
    Dim newfolder As String
    Dim newpath As String
    Dim fld As Object
    
    ' old file name
    filenm = "C:\Documents and Settings\user\Desktop\sample_1.xlsm"
   
   'new File Name
    newfolder = "C:\Documents and Settings\user\Desktop\ashish" ' please add "\" as the end
    
    ' new path
    ' add \ at the end of folder
    If VBA.Right(newfolder, 1) <> "\" Then newfolder = newfolder & "\"
    
    ' new path of file
    newpath = newfolder & VBA.Right(filenm, Len(filenm) - InStrRev(filenm, "\"))
    
    ' add some control check to avoid crashes
    
    ' check if file exists which we want to move
    If file_exists(filenm) <> "Exists" Then
        MsgBox "File does not exists so can not be moved ", vbInformation, "Note:"
        Exit Sub
    End If
    
    ' check if file already exits at destination folder
    If file_exists(newpath) = "Exists" Then
        MsgBox "File already exists at destination folder so can not be moved ", vbInformation, "Note:"
        Exit Sub
    End If

    ' check if  destination folder exists
    If folder_exists(newfolder) <> "Exists" Then
        MsgBox "Destination folder does not exists. Please create the folder first", vbInformation, "Note:"
        Exit Sub
    End If
    
    'move it finally

    Set fld = CreateObject("Scripting.FileSystemObject")
    fld.Movefile filenm, newfolder


End Sub
