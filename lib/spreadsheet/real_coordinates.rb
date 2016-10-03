# 2010.04.11

require 'spreadsheet'

require 'Module/override'

module OneBased
  module Spreadsheet
    module Worksheet
      
      def [](row, column)
        row(row - 1)[column - 1]
      end
      
      def []=(row, column, value)
        row(row - 1)[column - 1] = value
      end
      
    end
  end
end

Spreadsheet::Worksheet.send(:override, OneBased::Spreadsheet::Worksheet)

workbook = Spreadsheet.open('../../templates/3 - Request for Activations v1 - SPs.xls')
worksheet = workbook.worksheet(1)
worksheet[6,1] = '2010-04-11'
workbook.write('out.xls')
