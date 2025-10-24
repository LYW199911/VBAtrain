VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm2 
   Caption         =   "ファイル"
   ClientHeight    =   1344
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   5556
   OleObjectBlob   =   "UserForm2.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "UserForm2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public fileMap As Object

Private Sub UserForm_Initialize()
    Dim key As Variant
    If fileMap Is Nothing Then Exit Sub
    ComboBox1.Clear
    For Each key In fileMap.Keys
        ComboBox1.AddItem CStr(key)
    Next key
End Sub

Private Sub ComboBox1_Change()
    If fileMap Is Nothing Then Exit Sub
    Dim selectedFile As String
    selectedFile = ComboBox1.Value
    If selectedFile <> "" Then
        MsgBox "File path: " & vbCrLf & CStr(fileMap(selectedFile)), vbInformation
    End If
End Sub

Private Sub CommandButton1_Click()
    If ComboBox1.ListIndex < 0 Then
        MsgBox "Please select a file first.", vbExclamation
        Exit Sub
    End If
    Dim p As String
    p = CStr(fileMap(ComboBox1.Value))
    EnsureFolderExists Left$(p, InStrRev(p, "\") - 1)
    CreateExcelFile p
    MsgBox "File created successfully: " & p, vbInformation
End Sub

