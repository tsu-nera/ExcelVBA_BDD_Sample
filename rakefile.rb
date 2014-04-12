# -*- coding: utf-8 -*-
require 'rake/clean'
require 'win32ole'
require 'pp'

EXCEL_FILE  = "sample.xlsm"
DEBUG_SHOW = true
EXPORT_DIR_PATH = "./src"

task :default => "open"

desc "Open or Connect Excel File"
task :open do
  @book = openExcel(EXCEL_FILE)
end

desc "Save Excel File"
task :save => :open do
  @book.DisplayAlerts = false
  @book.Save
  @book.DisplayAlerts = true
end

desc "import All Modules"
task :import => :open do
  @book.run("importAllModules")
  @book.run("importFileManager")
end

desc "export all files to specified dir"
 task :export => :open do
  @book.run("exportAllModules")
  @book.run("exportFileManager")
end

# http://officetanaka.net/excel/vba/vbe/index.htm
desc "Open Visual Basic Editor for Application"
task :vbe => :open do
  @book.VBE.MainWindow.Visible = true
end

desc "Run All Tests"
task :spec => [:hide, :vbe, :import] do
  @book.run("RunAllTests")
end

desc "Make reliece excel file"
task :release do
  puts "to be implemented"
end

desc "Show Excel"
task :show => :open do
  @book.Visible = true
end

desc "Hide Excel"
task :hide => :open do
  @book.Visible = false
end

desc "Count Steps in Project"
task :step => [:hide, :vbe, :import] do
  @book.run("ShowTotalCodeLinesInProject")
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
