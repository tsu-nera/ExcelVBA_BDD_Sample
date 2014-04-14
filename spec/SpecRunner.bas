Attribute VB_Name = "SpecRunner"
'-------------------------------------------------------------------------------
' Name:
'-------------------------------------------------------------------------------
Public Sub RunAllTests()
  Calc_spec.Specs
  FileManager_spec.Specs
End Sub
