Attribute VB_Name = "FileManager_spec"
Sub Specs()
  On Error Resume Next
  Dim Specs As New SpecSuite
  Dim file_name As String
  Dim Result As Boolean

  With Specs.It("should not import .# file")
     file_name = ".#test.bas"
     Result = FileManager1.isValidImportFile(file_name)
    .Expect(Result).ToEqual False
  End With

  With Specs.It("should not import # file")
    file_name = "#test.bas#"
     Result = FileManager.isValidImportFile(file_name)
    .Expect(Result).ToEqual False
  End With

  With Specs.It("should not import *.cls~ file")
      file_name = "test.cls~"
     Result = FileManager.isValidImportFile(file_name)
    .Expect(Result).ToEqual False
  End With

  With Specs.It("should not import *.bas~ file")
     file_name = "test.bas~"
     Result = FileManager.isValidImportFile(file_name)
    .Expect(Result).ToEqual False
  End With
  
  With Specs.It("should import *.bas file")
     file_name = "test.bas"
     Result = FileManager.isValidImportFile(file_name)
    .Expect(Result).ToEqual True
  End With
  
  InlineRunner.RunSuite Specs
End Sub
