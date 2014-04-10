Attribute VB_Name = "FileManager1"
Option Explicit

Const INPORT_FOLDER As String = "src"

' モジュール定義
Enum Module
    Standard = 1   '標準モジュール
    Class = 2      'クラス モジュール
    Forms = 3      'Microsoft Forms
    ActiveX = 11   'ActiveX デザイナー
    Document = 100 'ドキュメント モジュール
End Enum

'-------------------------------------------------------------
' Name: reloadThisWorkbook
'
' http://social.msdn.microsoft.com/Forums/office/en-US/
' b823faa5-a432-4435-84cd-a04eadcbd1f5/
' loading-a-subroutine-into-thisworkbook-using-vba?forum=exceldev
'
' http://www.ozgrid.com/forum/showthread.php?t=26078
'-------------------------------------------------------------
Public Sub reloadThisWorkbook()
  Call clearThisWorkbook
  Call importThisWorkbook
End Sub

Private Sub clearThisWorkbook()
  With ThisWorkbook.VBProject.VBComponents("ThisWorkbook").CodeModule
    .DeleteLines StartLine:=1, count:=.CountOfLines
  End With
End Sub
 
Private Sub importThisWorkbook()
  Dim full_path As String
  full_path = ThisWorkbook.Path & "\" & INPORT_FOLDER & "\ThisWorkbook.cls"
  Debug.Print "Import from " & full_path

  With ThisWorkbook.VBProject.VBComponents("ThisWorkbook").CodeModule
     .AddFromFile full_path
     
     ' ゴミの除去
     .DeleteLines StartLine:=1, count:=4
  End With
End Sub

Private Sub clearModules()
  '標準モジュール/クラスモジュール初期化(全削除)
  
  Dim component As Object
  For Each component In ThisWorkbook.VBProject.VBComponents
      
    '標準モジュール/クラスモジュールを全て削除
    If component.Type = Module.Standard Or component.Type = Module.Class Then
      ThisWorkbook.VBProject.VBComponents.Remove component
    End If
    
  Next component
End Sub

Private Sub importAllModule()
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
            
             ' ゴミの除去
            .DeleteLines StartLine:=1, count:=4
          End With
        Case Else
          ThisWorkbook.VBProject.VBComponents.Import myFile
        End Select
      ElseIf myExtention = "bas" Then
          ThisWorkbook.VBProject.VBComponents.Import myFile
      End If
    Next myFile
    
    Set myFSO = Nothing
    Set myFolder = Nothing
    Set myFile = Nothing
End Sub

'-------------------------------------------------------------
' Name: ExportAllModule
' Func: すべてExport
' Reference from
' http://d.hatena.ne.jp/jamzz/20131002/1380696685
'-------------------------------------------------------------
Public Sub ExportAllModule()
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
        
        ' エクスポート
        full_path = ThisWorkbook.Path & "\" & export_path & "\" & vb_component.Name & extention
        Debug.Print "Export to " & full_path
        vb_component.Export full_path
    Next
End Sub

