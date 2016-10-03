# real_coordinates.rb

# 20161003
# 0.7.2

# Description: Wouldn't you rather use worksheet.a1 than worksheet[0,0]?

# Changes since 0.6:
# 1. + letter_places(), initially for use with column_letters(), but now I'm not so sure...
# 2. + place_values(), actually for use with column_letters().
# 3. ~ column_letters(), so that it actually works now!
# 4. ~ numeric_coordinates(), so that the parts are labelled.
# 5. + real_coordinates(), so as this is bidirectional for symmetry at least!
# 6. Far more extensive testing done.
# 0/1
# 7. - letter_places(), since it isn't being used.
# 8. Tidied the tests a little.
# 1/2
# 9. ~ column_letters(), removing some extraneous code.
# 10. Moved date and version information here.

require_relative 'RealCoordinates/Spreadsheet/Worksheet'
