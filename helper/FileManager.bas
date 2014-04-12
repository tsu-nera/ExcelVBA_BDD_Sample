Attribute VB_Name = "FileManager"
'-------------------------------------------------------------
' Name: FileManager
'-------------------------------------------------------------
Option Explicit

Public Const SRC_FOLDER As String = "src"       'source folder
Public Const SPEC_FOLDER As String = "spec"     'spec folder
Public Const HELPER_FOLDER As String = "helper" 'helper tool folder

Dim helper_files() As Variant
Const SPEC_SUFFIX As String = "_spec"

Enum Module
  Standard = 1
  Class = 2
  Forms = 3
  ActiveX = 11
  Document = 100
End Enum

Private Sub defineHelperFiles()
  helper_files = Array("FileManager", _
                       "FileManagerSub", _
                       "InlineRunner", _
                       "SpecDefinition", _
                       "SpecExpectation", _
                       "SpecRunner", _
                       "SpecSuite", _
                       "StepCounter", _
                       "mdlPrintF")
End Sub

'-------------------------------------------------------------
' Name: importAllModules()
' Func: Import All Modules  (without FileManager)
'-------------------------------------------------------------
Public Sub importAllModules()
  Call clearAllModules
  Call importAllModulesCore
End Sub

'-------------------------------------------------------------
' Name: release
' Func: Import Src Modules Only
'-------------------------------------------------------------
Public Sub release()
  Call clearAllModules
  Call importSrcMolues
End Sub

' It's dengerous procedure, be careful
Private Sub clearAllModules()
  Dim component As Object
  For Each component In ThisWorkbook.VBProject.VBComponents
      
    If component.Type = Module.Standard Or component.Type = Module.Class Then
      If Not isMySelf(component.name) Then
        ThisWorkbook.VBProject.VBComponents.Remove component
      End If
    End If
    
  Next component
End Sub

Private Sub importSrcMolues()
  Dim myFSO As New FileSystemObject
  Dim srcFolder As folder
  Dim specFolder As folder
  Dim helperFolder As folder
  
  Set srcFolder = myFSO.getFolder(ThisWorkbook.Path & "\" & SRC_FOLDER)
  Call importModulesIn(srcFolder.files)

  Set myFSO = Nothing
  Set srcFolder = Nothing
  Set specFolder = Nothing
  Set helperFolder = Nothing
End Sub

Private Sub importAllModulesCore()
  Dim myFSO As New FileSystemObject
  Dim srcFolder As folder
  Dim specFolder As folder
  Dim helperFolder As folder
  
  Set srcFolder = myFSO.getFolder(ThisWorkbook.Path & "\" & SRC_FOLDER)
  Call importModulesIn(srcFolder.files)
  
  Set specFolder = myFSO.getFolder(ThisWorkbook.Path & "\" & SPEC_FOLDER)
  Call importModulesIn(specFolder.files)
  
  Set helperFolder = myFSO.getFolder(ThisWorkbook.Path & "\" & HELPER_FOLDER)
  Call importModulesIn(helperFolder.files)
    
  Set myFSO = Nothing
  Set srcFolder = Nothing
  Set specFolder = Nothing
  Set helperFolder = Nothing
End Sub

Sub importModulesIn(files As files)
  Dim myFile As File

  For Each myFile In files
    
    If Not isValidImportFile(myFile.name) Then
      GoTo Next_myFile
    End If

    Debug.Print "Import from " & myFile

    If isExcelOnject(myFile.name) Then
      ' ThisWorkbook or Sheet_
      InsertLines (myFile)
    Else
      ' Standard or Class or Form Object
      ThisWorkbook.VBProject.VBComponents.Import myFile
    End If
    
Next_myFile:
  Next myFile
  
  Set myFile = Nothing
End Sub

' Excel Object is impossible to remove.
' Instead, delete all lines and insert.
Private Sub InsertLines(myFile As String)
  Dim myFSO As New FileSystemObject
  Dim myBaseName As String: myBaseName = myFSO.GetBaseName(myFile)
  
  With ThisWorkbook.VBProject.VBComponents(myBaseName).CodeModule
    .DeleteLines StartLine:=1, count:=.CountOfLines
    .AddFromFile myFile
    
    ' Delete header lines
    .DeleteLines StartLine:=1, count:=4
  End With

  Set myFSO = Nothing
End Sub

Public Function isExcelOnject(fileName As String) As Boolean
  Select Case Left(fileName, 5)
    Case "Sheet"
      isExcelOnject = True
    Case "ThisW"
      isExcelOnject = True
    Case Else
      isExcelOnject = False
  End Select
End Function

Public Function isValidImportFile(fileName As String) As Boolean
  Dim myFSO As FileSystemObject
  Set myFSO = CreateObject("Scripting.FileSystemObject")
  
  If Left(fileName, 1) = "." Then
    isValidImportFile = False
  ElseIf Left(fileName, 1) = "#" Then
    isValidImportFile = False
  ElseIf Right(fileName, 1) = "~" Then
        isValidImportFile = False
  ElseIf isMySelf(myFSO.GetBaseName(fileName)) Then
    isValidImportFile = False
  ElseIf Not hasValidExtention(fileName) Then
    isValidImportFile = False
  Else
    isValidImportFile = True
  End If

  Set myFSO = Nothing
End Function

Private Function isMySelf(baseName As String) As Boolean
  isMySelf = (baseName = "FileManager")
End Function

Private Function hasValidExtention(fileName As String) As Boolean
  Dim myFSO As FileSystemObject
  Set myFSO = CreateObject("Scripting.FileSystemObject")

  Select Case myFSO.GetExtensionName(fileName)
    Case "bas", "cls", "frm"
      hasValidExtention = True
    Case Else
      hasValidExtention = False
  End Select
  
  Set myFSO = Nothing
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
  Dim baseName As String
  Dim folder As String
  Dim vb_component As Object
    
  For Each vb_component In ThisWorkbook.VBProject.VBComponents
    
    baseName = vb_component.name
    extention = getExtention(vb_component.Type)
    folder = getFolder(vb_component.name)

    If Not isMySelf(baseName) Then
      full_path = getAbsolutePath(folder, baseName, extention)
      Debug.Print "Export to " & full_path
      vb_component.Export full_path
    End If
  Next
End Sub

Private Function getExtention(vbCompType As Integer) As String
  Dim extention As String
  
  Select Case vbCompType
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

Private Function getFolder(baseName As String) As String
  
  If isHelperFile(baseName) Then
    getFolder = HELPER_FOLDER
  ElseIf isSpecFile(baseName) Then
    getFolder = SPEC_FOLDER
  Else
    getFolder = SRC_FOLDER
  End If
End Function

Public Function isSrcFile(baseName As String) As Boolean
  If isSpecFile(baseName) Then
    isSrcFile = False
  ElseIf isHelperFile(baseName) Then
    isSrcFile = False
  Else
    isSrcFile = True
  End If
End Function

Private Function isSpecFile(baseName As String) As Boolean
  If Right(baseName, 5) = SPEC_SUFFIX Then
    isSpecFile = True
  Else
    isSpecFile = False
  End If
End Function

Private Function isHelperFile(baseName As String) As Boolean
  Dim fileName As Variant
  Call defineHelperFiles

  For Each fileName In helper_files
    If fileName = baseName Then
      isHelperFile = True
      Exit Function
    End If
  Next
          
  isHelperFile = False
End Function

Private Function getAbsolutePath(folder As String, _
                                  baseName As String, _
                                  extName As String) As String
  getAbsolutePath = ThisWorkbook.Path & "\" & folder & "\" & baseName & extName
End Function
