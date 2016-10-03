# test.rb

# 20121130, 1204
# 0.5.0

require File.expand_path(File.join(File.dirname(__FILE__), '..', 'lib', 'spreadsheet', 'real_coordinates'))

workbook = Spreadsheet.open('test.xls')
worksheet = workbook.worksheet(0)

%w{a b c d e}.each do |column|
  (1..5).each do |row|
    worksheet.send("#{column}#{row}=", "#{column}, #{row}")
  end
end
workbook.write('out.xls')
