VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "メニュー"
   ClientHeight    =   2532
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   2580
   OleObjectBlob   =   "UserForm1.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub UserForm_Initialize()
    Call createFolderPath
End Sub

' Create folder
Private Sub CommandButton1_Click()
    Call createFolderPath
    EnsureFolderExists folderPath
    MsgBox "Folder created or verified successfully." & vbCrLf & folderPath, vbInformation
End Sub

' Read filenames and build dictionary (from 2nd row to right)
Private Sub CommandButton2_Click()
    Dim i As Long
    Call createFolderPath

    ' Ensure folder exists
    EnsureFolderExists folderPath

    If Len(Dir(folderPath, vbDirectory)) = 0 Then
        MsgBox "Folder does not exist. Please create it first.", vbExclamation
        Exit Sub
    End If

    Set fileDict = CreateObject("Scripting.Dictionary")
    i = 2

    Do While Trim$(WS().Cells(2, i).Value) <> ""
        createFilePath i
        If Not fileDict.Exists(fileName) Then
            fileDict.Add fileName, filePath
            ' Actually create the file
            CreateExcelFile filePath
        End If
        i = i + 1
    Loop

    MsgBox "All files have been created successfully!" & vbCrLf & _
           "Total files: " & fileDict.Count, vbInformation
End Sub

' Open UserForm2 directly (requires dictionary already loaded)
Private Sub CommandButton3_Click()
    If fileDict Is Nothing Then
        MsgBox "Please load the file list first.", vbExclamation
        Exit Sub
    End If
    Dim uf2 As UserForm2
    Set uf2 = New UserForm2
    Set uf2.fileMap = fileDict
    uf2.Show vbModal
End Sub

