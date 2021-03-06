Attribute VB_Name = "FileManagerSub"
'-------------------------------------------------------------
' Name: FileManagerSub
'-------------------------------------------------------------
Option Explicit
Const FILEMANAGER_PATH As String = "helper\FileManager.bas"

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
' Name: clearFileManager
' Func: clear FileManager.bas
'-------------------------------------------------------------------------------
Public Sub clearFileManager()
  If ModExists("FileManager") Then
    ThisWorkbook.VBProject.VBComponents.Remove _
      ThisWorkbook.VBProject.VBComponents("FileManager")
  End If
End Sub

'-------------------------------------------------------------------------------
' Name: importFileManager
' Func: Import FileManager.bas
'-------------------------------------------------------------------------------
Public Sub importFileManager()
  clearFileManager
    
  Debug.Print "Import from " & getAbsoluteFileManagerPath
  ThisWorkbook.VBProject.VBComponents.Import getAbsoluteFileManagerPath
End Sub

' Check if Module Exist
' http://forums.arcgis.com/threads/5601-How-to-check-if-a-module-exist
Function ModExists(name As String) As Boolean

  ModExists = False
  Dim pVBE As VBIDE.VBE
  Set pVBE = Application.VBE
  Dim l As Long
  For l = 1 To pVBE.VBProjects.count
    Dim k As Long
    For k = 1 To pVBE.VBProjects(l).VBComponents.count
      If pVBE.VBProjects(l).VBComponents(k).Type = vbext_ct_StdModule Then
        Dim s As String
        s = UCase(pVBE.VBProjects.Item(l).VBComponents(k).name)
        If s = UCase(name) Then
          ModExists = True
          Exit Function
        End If
      End If
    Next k
  Next l
End Function

Private Function getAbsoluteFileManagerPath() As String
  getAbsoluteFileManagerPath = ThisWorkbook.Path & "\" & FILEMANAGER_PATH
End Function
