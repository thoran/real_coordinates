# RealCoordinates/Spreadsheet/Worksheet.rb
# RealCoordinates/Spreadsheet/Worksheet

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

      def place_values(column_number)
        result = []
        counter = 0
        until counter > column_number do
          place = 0
          loop do
            if result[place].nil?
              result[place] = 0
              break
            elsif result[place] < 25
              result[place] += 1
              break
            else
              result[place] = 0
              place += 1
            end
          end
          counter += 1
        end
        result.reverse
      end

      def column_letters(column_number)
        letters = ('A'..'Z').to_a
        place_values(column_number).collect do |place_value|
          letters[place_value]
        end.join('')
      end

      def numeric_coordinates(letters_and_numbers_coordinates)
        parts = letters_and_numbers_coordinates.match(/^(\w+?)(\d+?)$/)
        row_number, column_number = parts[2].to_i - 1, column_number(parts[1])
        [row_number, column_number]
      end

      def real_coordinates(*numeric_coordinates)
        numeric_coordinates = numeric_coordinates.flatten
        row_number, column_number = numeric_coordinates.first.to_i, numeric_coordinates.last.to_i
        "#{column_letters(column_number)}#{row_number + 1}"
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
