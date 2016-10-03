# RealCoordinates

# 20121204, 05
# 0.6.0

# Description: Wouldn't you rather use worksheet.a1 than worksheet[0,0]?  

# Changes since 0.5: 
# 1. Now using Minitest for testing.  
# 2. + column_letters().  

require 'spreadsheet'

module RealCoordinates
  module Spreadsheet
    module Worksheet

      def column_number(column_letters)
        column_letters = column_letters.upcase
        letters = ('A'..'Z').to_a
        place_value = 26**(column_letters.size)
        column_number = -1
        column_letters.each_char do |letter|
          column_number += (letters.index(letter) + 1) * (place_value /= 26)
        end
        column_number
      end

      def column_letters(column_number)
        powers = 1
        until column_number.to_f/26**(powers) < 1
          powers += 1
        end
        powers = (powers <= 0 ? 0 : powers -= 1)
        column_letters = ''
        letters = ('A'..'Z').to_a
        residual_value = column_number
        powers.downto(0) do |power|
          p power
          p residual_value.divmod(26**power)
          index_value, residual_value = residual_value.divmod(26**power)
          column_letters << letters[index_value]
          puts
        end
        column_letters
      end

      def numeric_coordinates(letters_and_numbers_coordinates)
        parts = letters_and_numbers_coordinates.match(/^(\w+?)(\d+?)$/)
        [parts[2].to_i - 1, column_number(parts[1])]
      end

      def method_missing(method_name, *args, &block)
        if method_name.to_s =~ /=$/
          coords = numeric_coordinates(method_name.to_s.sub(/=$/, ''))
          self.send('[]=', coords[0], coords[1], *args)
        else
          coords = numeric_coordinates(method_name.to_s)
          self.send('[]', coords[0], coords[1])
        end
      end

    end
  end
end

Spreadsheet::Worksheet.send(:include, RealCoordinates::Spreadsheet::Worksheet)
