require 'spreadsheet'

# create a new book and sheet
book = Spreadsheet::Workbook.new
sheet = book.create_worksheet
5.times {|j| 5.times {|i| sheet[j,i] = (i+1)*10**j}}

# column
sheet.column(2).hidden = true
sheet.column(3).hidden = true
sheet.column(2).outline_level = 1
sheet.column(3).outline_level = 1

# row
sheet.row(2).hidden = true
sheet.row(3).hidden = true
sheet.row(2).outline_level = 1
sheet.row(3).outline_level = 1

# save file
book.write 'out.xls'
