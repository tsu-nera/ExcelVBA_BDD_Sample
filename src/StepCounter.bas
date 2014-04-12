Attribute VB_Name = "StepCounter"
'-----------------------------------------------------------------------------------
' Name: StepCounte
'   Attention:
'   Need  Microsoft Visual Basic for Applications Extensibility
' http://www.cpearson.com/excel/vbe.aspx
' http://excelappwithvba.web.fc2.com/generating_report_sheet/attaching_vba_code.html
'-----------------------------------------------------------------------------------
Option Explicit

' TODO Define IgnoreList

'---------------------------------------------------------------------------------
' Name: ShowTotalCodeLinesInProject
'----------------------------------------------------------------------------------
Public Sub ShowTotalCodeLinesInProject()
  Dim vbcComp As VBIDE.VBComponent
  Dim vbcLine As Integer

  For Each vbcComp In Application.VBE.ActiveVBProject.VBComponents
    vbcLine = TotalCodeLinesInVBComponent(vbcComp)
    Debug.Print vbcCode.name & "   " & vbcLine
  Next vbcComp
End Sub

Public Sub ExportTotalCodeLinesInProject()
End Sub

Public Sub ExportTotalCodeLinesInProjectToCSV()
  Dim vbcComp As VBIDE.VBComponent
  Dim vbcLine As Integer

  For Each vbcComp In Application.VBE.ActiveVBProject.VBComponents
    vbcLine = TotalCodeLinesInVBComponent(vbcComp)
    Debug.Print vbcLine
  Next vbcComp
End Sub

'---------------------------------------------------------------------------------
' Name: TotalCodeLinesInProject
'----------------------------------------------------------------------------------
Private Function TotalCodeLinesInProject(VBProj As VBIDE.VBProject) As Long
 
  Dim VBComp As VBIDE.VBComponent
  Dim LineCount As Long

  If VBProj.Protection = vbext_pp_locked Then
    TotalCodeLinesInProject = -1
    Exit Function
  End If

  For Each VBComp In VBProj.VBComponents
    LineCount = LineCount + TotalCodeLinesInVBComponent(VBComp)
  Next VBComp
  
  TotalCodeLinesInProject = LineCount
End Function

'---------------------------------------------------------------------------------
' Name: TotalCodeLinesInProject
'----------------------------------------------------------------------------------
Private Function TotalCodeLinesInVBComponent(VBComp As VBIDE.VBComponent) As Long
  Dim N As Long
  Dim s As String
  Dim LineCount As Long
  
  If VBComp.Collection.Parent.Protection = vbext_pp_locked Then
    TotalCodeLinesInVBComponent = -1
    Exit Function
  End If
  
  With VBComp.CodeModule
    For N = 1 To .CountOfLines
      s = .Lines(N, 1)
      If Trim(s) = vbNullString Then
        ' blank line, skip it
      ElseIf Left(Trim(s), 1) = "'" Then
        ' comment line, skip it
      Else
        LineCount = LineCount + 1
      End If
    Next N
  End With
  TotalCodeLinesInVBComponent = LineCount
End Function

