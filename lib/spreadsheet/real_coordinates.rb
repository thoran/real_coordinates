# RealCoordinates

# 20121130, 1204
# 0.5.0

# Description: Wouldn't you rather use worksheet.a1 than worksheet[0,0]?  

# Changes since 2010.04.11 (0.6): 
# 1. - #[].  
# 2. - #[]=.  
# 3. - require 'Module', since #[] and #[]= are now gone and I don't need to override methods anymore.  
# 4. ~ column_number(), so that it no longer defaults to being one-based, and no longer optionally can be zero-based, but defaults to being zero-based.  
# 5. - column_number(), alias_method :one_based_column_number, :column_number, since this is zero-based only now.  
# 6. - column_number(), alias_method :one_based_numeric_column_number, :column_number, since this is zero-based only now.  
# 7. - zero_based_column_number(), since column_number is now zero-based.  
# 8. ~ numeric_coordinates(), so that it no longer defaults to being one-based, and no longer optionally can be zero-based, but defaults to being zero-based.  
# 9. - numeric_coordinates(), alias_method :one_based_coordinates, :numeric_coordinates, since this is zero-based only now.  
# 9. - numeric_coordinates(), alias_method :one_based_numeric_coordinates, :numeric_coordinates, since this is zero-based only now.  
# 10. - zero_based_coordinates(), since numeric_coordinates is now zero-based.  
# 11. ~ method_missing(), so that it now uses zero_based_coordinates().  

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
