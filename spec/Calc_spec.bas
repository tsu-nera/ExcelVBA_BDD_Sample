Attribute VB_Name = "Calc_spec"
'-------------------------------------------------------------------
' Name     : Calc_spec
'-------------------------------------------------------------------
Sub Specs()
        On Error Resume Next
        Dim Specs As New SpecSuite

        Dim Subject As New Calc

        With Specs.It("should add two numbers")
                ' Test the desired behavior
                .Expect(Subject.Add(2, 2)).ToEqual 4
                .Expect(Subject.Add(3, -1)).ToEqual 2
                .Expect(Subject.Add(-1, -2)).ToEqual -3
        End With

        InlineRunner.RunSuite Specs
End Sub
