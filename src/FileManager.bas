Attribute VB_Name = "FileManager"
'-------------------------------------------------------------
' Name: FileManager
'-------------------------------------------------------------
Option Explicit

Public Const WORK_FOLDER As String = "src" 'Your Working Directory

Enum Module
  Standard = 1
  Class = 2
  Forms = 3
  ActiveX = 11
  Document = 100
End Enum

'-------------------------------------------------------------
' Name: importAllModules()
' Func: Import All Modules  (without FileManager)
'-------------------------------------------------------------
Public Sub importAllModules()
  Dim myFSO As New FileSystemObject
  Dim myFolder As Folder
  Dim myFile As File
  Dim myExtention As String
  Dim myBaseName As String
  
  Set myFSO = CreateObject("Scripting.FileSystemObject")
  Set myFolder = myFSO.GetFolder(ThisWorkbook.Path & "\" & WORK_FOLDER)

  Call clearModules
    
  For Each myFile In myFolder.Files
    myExtention = myFSO.GetExtensionName(myFile.name)
    myBaseName = myFSO.GetBaseName(myFile.name)
    
    If Not isValidImportFile(myFile.name) Then
      GoTo Next_myFile
    End If

    If myExtention = "cls" Then
      Select Case Left(myBaseName, 5)
        Case "Sheet", "ThisW"
          With ThisWorkbook.VBProject.VBComponents(myBaseName).CodeModule
            .DeleteLines StartLine:=1, count:=.CountOfLines
            .AddFromFile myFile

            ' Delete header lines
            .DeleteLines StartLine:=1, count:=4
          End With
        Case Else
          ThisWorkbook.VBProject.VBComponents.Import myFile
      End Select
    ElseIf myExtention = "bas" Then
      ThisWorkbook.VBProject.VBComponents.Import myFile
    End If
    
Next_myFile:
  Next myFile
    
  Set myFSO = Nothing
  Set myFolder = Nothing
  Set myFile = Nothing
End Sub

' It's dengerous procedure, be careful
Private Sub clearModules()
  Dim component As Object
  For Each component In ThisWorkbook.VBProject.VBComponents
      
    If component.Type = Module.Standard Or component.Type = Module.Class Then
      If Not Myself(component.name) Then
        ThisWorkbook.VBProject.VBComponents.Remove component
      End If
    End If
    
  Next component
End Sub

Public Function isValidImportFile(filename As String) As Boolean

  Dim myFSO As New FileSystemObject
  Set myFSO = CreateObject("Scripting.FileSystemObject")
  
  If Left(filename, 1) = "." Then
    isValidImportFile = False
  ElseIf Left(filename, 1) = "#" Then
    isValidImportFile = False
  ElseIf Right(filename, 1) = "~" Then
        isValidImportFile = False
  ElseIf Myself(myFSO.GetBaseName(filename)) Then
    isValidImportFile = False
  Else
    isValidImportFile = True
  End If

  Set myFSO = Nothing
End Function

Private Function Myself(baseName As String) As Boolean
  Myself = (baseName = "FileManager")
End Function

'-------------------------------------------------------------
' Name: exportAllModules
' Func: Export All Files (without FileManager)
' Reference from
' http://d.hatena.ne.jp/jamzz/20131002/1380696685
'-------------------------------------------------------------
Public Sub exportAllModules()
  Dim full_path As String
  Dim extention As String
  Dim vb_component As Object
    
  For Each vb_component In ThisWorkbook.VBProject.VBComponents
    
    extention = getExtention(vb_component)

    If Not Myself(vb_component.name) Then
      full_path = getAbsolutePath(vb_component.name, extention)
      Debug.Print "Export to " & full_path
      vb_component.Export full_path
    End If
  Next
End Sub

Private Function getExtention(myComponent As VBComponent) As String
  Dim extention As String
  
  Select Case myComponent.Type
    Case Module.Standard
      extention = ".bas"
    Case Module.Class
      extention = ".cls"
    Case Module.Forms
      extention = ".frm"
    Case Module.ActiveX
      extention = ".cls"
    Case Module.Document
      extention = ".cls"
  End Select

  getExtention = extention
End Function

Private Function getAbsolutePath(baseName, extName) As String
  getAbsolutePath = ThisWorkbook.Path & "\" & WORK_FOLDER & "\" & baseName & extName
End Function
