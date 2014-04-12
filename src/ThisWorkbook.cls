VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'-------------------------------------------------------------------------------
' Name: ThisWorkBook
'-------------------------------------------------------------------------------
Option Explicit
Const EXPORT_FOLDER As String = "src"
Const FILEMANAGER_PATH As String = "src\FileManager.bas"

'-------------------------------------------------------------------------------
' Name: exportFileManager
' Func:  Export FileManager.bas
'-------------------------------------------------------------------------------
Public Sub exportFileManager()
  Dim myFileManager As VBComponent
  Set myFileManager = ThisWorkbook.VBProject.VBComponents("FileManager")

  Debug.Print "Export to " & getAbsoluteFileManagerPath
  myFileManager.Export getAbsoluteFileManagerPath

  Set myFileManager = Nothing
End Sub

'-------------------------------------------------------------------------------
' Name: importFileManager
' Func: Import FileManager.bas
'-------------------------------------------------------------------------------
Public Sub importFileManager()
  ThisWorkbook.VBProject.VBComponents.Remove _
    ThisWorkbook.VBProject.VBComponents("FileManager")
    
  ThisWorkbook.VBProject.VBComponents.Import getAbsoluteFileManagerPath
End Sub

Private Function getAbsoluteFileManagerPath() As String
  getAbsoluteFileManagerPath = ThisWorkbook.Path & "\" & FILEMANAGER_PATH
End Function


