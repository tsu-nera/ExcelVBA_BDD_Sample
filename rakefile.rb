# -*- coding: utf-8 -*-
require 'rake/clean'
require 'win32ole'

EXCEL_FILE  = "sample.xlsm"
DEBUG_SHOW = true  

task :default => "open"

desc "Open or Connect Excel File"
task :open do
  @book = openExcel(EXCEL_FILE)
end

desc "import All Modules"
task :import => :open do
  @book.run("ThisWorkBook.reloadModule")
end

desc "export all files to specified dir"
task :export => :open do
  @book.run("ThisWorkBook.ExportAll")
end

desc "Run All Tests"
# task :spec => [:open, :reload] do
task :spec => :open do
  @book.run("ThisWorkBook.RunAllTests")
end

desc "make reliece excel file"
task :release do
  puts "to be implemented"
end

# refered from 
# http://osdir.com/ml/lang.ruby.japanese/2005-11/msg00180.html
def getAbsolutePath filename
  fso = WIN32OLE.new('Scripting.FileSystemObject')
  return fso.GetAbsolutePathName(filename)
end

def openExcel(filename)
  filename = getAbsolutePath(filename)
  book = nil
  begin
    book = WIN32OLE::connect("Excel.Application")
  rescue WIN32OLERuntimeError
    book = WIN32OLE.new("Excel.Application")
  end
  book.Workbooks.each do |sheet|
    if sheet.FullName == filename
      sheet.Activate
    end
  end

  unless book.ActiveWorkbook && book.ActiveWorkbook.FullName == filename
    book.Workbooks.Open(filename)
  end
  book.Visible = DEBUG_SHOW
  return book
end

# refered from 
# http://d.hatena.ne.jp/jamzz/20131002/1380696685
# def export()
#   excel = WIN32OLE.new('Excel.Application')
#   excel_file = getAbsolutePath(EXCEL_FILE)
#   book = excel.Workbooks.Open(excel_file)
  
#   book.VBProject.VBComponents.each do |vb_component|

#     full_path = getExportPath(vb_component)

#     # export
#     p "export to " + full_path
#     # it doesn't work
#     # vb_component.Export full_path
#   end
# end

# # 標準モジュール
# Const_vbext_ct_StdModule = 1
# # クラス モジュール
# Const_vbext_ct_ClassModule = 2
# # Microsoft Forms
# Const_vbext_ct_MSForm = 3
# # ActiveX デザイナー
# Const_vbext_ct_ActiveXDesigner = 11
# # ドキュメント モジュール
# Const_vbext_ct_Document = 100

# def getExportPath(vb_component)
#   case vb_component.Type
#   when Const_vbext_ct_StdModule
#     extention = '.bas'
#   when Const_vbext_ct_ClassModule
#     extention = '.cls'
#   when Const_vbext_ct_MSForm
#     extention = '.frm'
#   when Const_vbext_ct_ActiveXDesigner
#     extention = '.cls'
#   when Const_vbext_ct_Document
#     extention = '.cls'
#   end

#   # absolute path
#   export_path = File.expand_path(EXPORT_DIR_PATH)

#   # full path
#   full_path =  File.join(export_path, vb_component.Name + extention)

#   return full_path
# end
