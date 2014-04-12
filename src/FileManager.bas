Attribute VB_Name = "FileManager"

Option Explicit

Const INPORT_FOLDER As String = "src" 'Export Directory

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
    Set myFolder = myFSO.GetFolder(ThisWorkbook.Path & "\" & INPORT_FOLDER)

    Call clearModules
    
    For Each myFile In myFolder.Files
      myExtention = myFSO.GetExtensionName(myFile.Name)
      myBaseName = myFSO.GetBaseName(myFile.Name)
      
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
      ElseIf myBaseName = "FileManager" Then
        'Nop
      ElseIf myExtention = "bas" Then
          ThisWorkbook.VBProject.VBComponents.Import myFile
      End If
    Next myFile
    
    Set myFSO = Nothing
    Set myFolder = Nothing
    Set myFile = Nothing
End Sub

Private Sub clearModules()
  
  Dim component As Object
  For Each component In ThisWorkbook.VBProject.VBComponents
      
    If component.Type = Module.Standard Or component.Type = Module.Class Then
      If component.Name <> "FileManager" Then
        ThisWorkbook.VBProject.VBComponents.Remove component
      End If
    End If
    
  Next component
End Sub

Public Function isValidImportFile(filename As String) As Boolean
  isValidImportFile = True
End Function

'-------------------------------------------------------------
' Name: exportAllModules
' Func: Export All Files (without FileManager)
' Reference from
' http://d.hatena.ne.jp/jamzz/20131002/1380696685
'-------------------------------------------------------------
Public Sub exportAllModules()
    Dim export_path As String
    Dim full_path As String
    Dim vb_component As Object
    
    export_path = INPORT_FOLDER
    
    For Each vb_component In ThisWorkbook.VBProject.VBComponents

        Dim extention As String
        Select Case vb_component.Type
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
        
        full_path = ThisWorkbook.Path & "\" & export_path & "\" & vb_component.Name & extention

        If vb_component.Name <> "FileManager" Then
          Debug.Print "Export to " & full_path
          vb_component.Export full_path
        End If
    Next
End Sub

