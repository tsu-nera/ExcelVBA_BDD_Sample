Attribute VB_Name = "ThisWorkbook_spec"
'-------------------------------------------------------------------
' Name     : Calc_spec
'-------------------------------------------------------------------
Sub Specs()
        On Error Resume Next
        Dim Specs As New SpecSuite

        InlineRunner.RunSuite Specs
End Sub
