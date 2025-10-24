Attribute VB_Name = "Module1"
Option Explicit

Public folderName As String
Public folderPath As String
Public fileName As String
Public filePath As String
Public fileDict As Object
Public iniPath As String

Public Function WS() As Worksheet
    ' Change the sheet index or name if needed
    Set WS = ThisWorkbook.Worksheets(1)
End Function

Sub menuButton_Click()
    UserForm1.Show vbModeless
End Sub

' Construct the folder path
Public Sub createFolderPath()
    folderName = Trim(WS().Cells(1, 2).Value)
    iniPath = Trim(WS().Cells(3, 2).Value)

    If folderName = "" Or iniPath = "" Then
        MsgBox "Folder name or base path is missing.", vbExclamation
        Exit Sub
    End If

    If Right$(iniPath, 1) <> "\" Then iniPath = iniPath & "\"
    folderPath = iniPath & folderName
End Sub

' Combine file path (reads filename from column i)
Public Sub createFilePath(ByVal i As Long)
    Call createFolderPath
    fileName = Trim(WS().Cells(2, i).Value)
    If fileName = "" Then
        MsgBox "File name is missing.", vbExclamation
        Exit Sub
    End If
    filePath = folderPath & "\" & fileName
End Sub

' Recursively create folder if not exists
Public Sub EnsureFolderExists(ByVal path As String)
    If Len(Dir(path, vbDirectory)) = 0 Then
        Dim parent As String
        parent = Left$(path, InStrRev(path, "\") - 1)
        If Len(parent) > 0 And Len(Dir(parent, vbDirectory)) = 0 Then
            EnsureFolderExists parent
        End If
        MkDir path
    End If
End Sub

' Get list of files inside folder
Public Function GetFileList() As Collection
    Dim files As New Collection
    Dim fName As String
    Call createFolderPath
    If Len(Dir(folderPath, vbDirectory)) = 0 Then
        Set GetFileList = files
        Exit Function
    End If
    fName = Dir(folderPath & "\*.*")
    Do While fName <> ""
        files.Add fName
        fName = Dir()
    Loop
    Set GetFileList = files
End Function

' Create a new Excel file (in same process)
Public Sub CreateExcelFile(ByVal outPath As String)
    Dim wb As Workbook
    Set wb = Application.Workbooks.Add
    Application.DisplayAlerts = False
    wb.SaveAs fileName:=outPath
    Application.DisplayAlerts = True
    wb.Close SaveChanges:=False
End Sub

