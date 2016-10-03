# 2010.04.11

require 'spreadsheet'

require 'Module/override'

module RealCoordinates
  module Spreadsheet
    module Worksheet
      
      def [](row, column)
        row(row - 1)[column - 1]
      end
      
      def []=(row, column, value)
        row(row - 1)[column - 1] = value
      end
      
      def column_number(column_letters, zero_based = false)
        column_letters = column_letters.upcase
        letters = ('A'..'Z').to_a
        place_value = 26**(column_letters.size)
        column_number = 0
        column_letters.each_char do |letter|
          column_number += (letters.index(letter) + 1) * (place_value /= 26)
        end
        if zero_based
          column_number - 1
        else
          column_number
        end
      end
      alias_method :one_based_column_number, :column_number
      alias_method :one_based_numeric_column_number, :column_number
      
      def zero_based_column_number(column_letters)
        column_number(column_letters, true)
      end
      
      def numeric_coordinates(letters_and_numbers_coordinates, zero_based = false)
        parts = letters_and_numbers_coordinates.match(/^(\w+?)(\d+?)$/)
        if zero_based
          [parts[2].to_i - 1, zero_based_column_number(parts[1])]
        else
          [parts[2].to_i, column_number(parts[1])]
        end
      end
      alias_method :one_based_coordinates, :numeric_coordinates
      alias_method :one_based_numeric_coordinates, :numeric_coordinates
      
      def zero_based_coordinates(letters_and_numbers_coordinates)
        numeric_coordinates(letters_and_numbers_coordinates, true)
      end
      alias_method :zero_based_numeric_coordinates, :zero_based_coordinates
      
      def method_missing(method_name, *args, &block)
        if method_name.to_s =~ /=$/
          coords = one_based_coordinates(method_name.to_s.sub(/=$/, ''))
          self.send('[]=', coords[0], coords[1], *args)
        else
          coords = one_based_coordinates(method_name.to_s)
          self.send('[]', coords[0], coords[1])
        end
      end
      
    end
  end
end

Spreadsheet::Worksheet.send(:override, RealCoordinates::Spreadsheet::Worksheet)

workbook = Spreadsheet.open('../../templates/3 - Request for Activations v1 - SPs.xls')
worksheet = workbook.worksheet(1)
worksheet.a6 = '2010-04-11'
workbook.write('out.xls')
