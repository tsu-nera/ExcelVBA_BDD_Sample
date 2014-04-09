VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Calc_spec"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'-------------------------------------------------------------------
' Name     : Calc_spec
'-------------------------------------------------------------------
Sub Specs()
        ' Create a new collection of specs
        On Error Resume Next
        Dim Specs As New SpecSuite

        Dim subject As New Calc

        ' Describe the desired behavior
        With Specs.It("should add two numbers")
                ' Test the desired behavior
                .Expect(subject.Add(2, 2)).ToEqual 4
                .Expect(subject.Add(3, -1)).ToEqual 2
                .Expect(subject.Add(-1, -2)).ToEqual -3
        End With

        ' Run the specs inline (in the Immediate window)
        InlineRunner.RunSuite Specs
End Sub
