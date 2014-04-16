Attribute VB_Name = "StepCounter"
'-----------------------------------------------------------------------------------
' Name: StepCounte
'   Attention:
'   Need  Microsoft Visual Basic for Applications Extensibility
' http://www.cpearson.com/excel/vbe.aspx
' http://excelappwithvba.web.fc2.com/generating_report_sheet/attaching_vba_code.html
'-----------------------------------------------------------------------------------
Option Explicit

Const COUNT_TXT_FILENAME As String = "count.txt"
Const COUNT_CSV_FILENAME As String = "count.csv"

'---------------------------------------------------------------------------------
' Name: ShowTotalCodeLinesInProject
'----------------------------------------------------------------------------------
Public Sub ShowTotalCodeLinesInProject()
  Dim vbcComp As VBIDE.VBComponent
  Dim vbcLine As Integer
  Dim TotalCount As Long: TotalCount = 0
  
  Dim str As String: str = ""
  
  str = str + SPrintF("-----------------------\n")
  str = str + SPrintF(" FileName      Execute \n")
  str = str + SPrintF("-----------------------\n")

  For Each vbcComp In Application.VBE.ActiveVBProject.VBComponents
    If FileManager.isSrcFile(vbcComp.name) Then
      vbcLine = TotalCodeLinesInVBComponent(vbcComp)
      TotalCount = TotalCount + vbcLine
      str = str + SPrintF(" %-17s%4d \n", vbcComp.name, vbcLine)
    End If
  Next vbcComp
  
  str = str + SPrintF("-----------------------\n")
  str = str + SPrintF(" Sum              %4d \n", TotalCount)
  str = str + SPrintF("-----------------------\n")

  Debug.Print str
End Sub

'---------------------------------------------------------------------------------
' Name: WriteTotalCodeLinesInProjectToTxt
'----------------------------------------------------------------------------------
Public Sub WriteTotalCodeLinesInProjectToTxt()
  Dim FSO As New FileSystemObject
  Dim outputFile As String
  
  outputFile = ThisWorkbook.Path & "\" & COUNT_TXT_FILENAME
  With FSO.OpenTextFile(fileName:=outputFile, _
                        IOMode:=ForWriting, Create:=True)

    .WriteLine ("-----------------------")
    .WriteLine (" FileName      Execute ")
    .WriteLine ("-----------------------")

    Dim vbcComp As VBIDE.VBComponent
    Dim vbcLine As Integer
    Dim TotalCount As Long: TotalCount = 0
      
     For Each vbcComp In Application.VBE.ActiveVBProject.VBComponents
       If FileManager.isSrcFile(vbcComp.name) Then
	 vbcLine = TotalCodeLinesInVBComponent(vbcComp)
	 TotalCount = TotalCount + vbcLine
	 .WriteLine(SPrintF(" %-17s%4d ", vbcComp.name, vbcLine))
       End If
     Next vbcComp
 
     .WriteLine ("-----------------------")
     .WriteLine(SPrintF(" Sum              %4d ", TotalCount))
     .WriteLine ("-----------------------")
    .Close
  End With
End Sub

'---------------------------------------------------------------------------------
' Name: TotalCodeLinesInProject
'----------------------------------------------------------------------------------
Private Sub TotalCodeLinesInProject(mode As Integer)

'
'  str = str + SPrintF("-----------------------\n")
'  str = str + SPrintF(" FileName      Execute \n")
'  str = str + SPrintF("-----------------------\n")
'
'  For Each vbcComp In Application.VBE.ActiveVBProject.VBComponents
'    If FileManager.isSrcFile(vbcComp.name) Then
'      vbcLine = TotalCodeLinesInVBComponent(vbcComp)
'      TotalCount = TotalCount + vbcLine
'      str = str + SPrintF(" %-17s%4d \n", vbcComp.name, vbcLine)
'    End If
'  Next vbcComp
'
'  str = str + SPrintF("-----------------------\n")
'  str = str + SPrintF(" Sum              %4d \n", TotalCount)
'  str = str + SPrintF("-----------------------\n")
'
'  Debug.Print str
End Sub

' Private Sub TotalCodeLineInProject(mode As Integer)
' End Sub

'---------------------------------------------------------------------------------
' Name: TotalCodeLinesInCOmmponent
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
