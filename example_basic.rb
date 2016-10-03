# 2010.04.11

########################################################################
# example_basic.rb
# 
# Simply run this and then open up basic_example.xls with MS Excel
# (or Gnumeric or whatever) to view the results.
########################################################################
base = File.basename(Dir.pwd)
if base == "examples" || base =~ /spreadsheet/i
   Dir.chdir("..") if base == "examples"
   $LOAD_PATH.unshift(Dir.pwd)
   $LOAD_PATH.unshift(Dir.pwd + "/lib")
   Dir.chdir("examples") rescue nil
end

require "spreadsheet/excel"
include Spreadsheet

puts "VERSION: " + Excel::VERSION

workbook = Excel.new("basic_example1.xls")

worksheet = workbook.add_worksheet

class Worksheet
  
  # 1-based placement, rather than 0-based placement
  def place(row, col, data = nil, format = nil)
    write(row - 1, col - 1, data, format)
  end
  
end

format3 = Format.new{ |f|
   f.color = "red"
   f.bold  = true
   f.underline = true
}
workbook.add_format(format3)

worksheet.place(1, 1, "Hello")
worksheet.place(2, 1, ["Matz","Larry","Guido"])

workbook.close
