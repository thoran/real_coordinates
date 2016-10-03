# test.rb

# 20121204
# 0.6.0

# Notes: 
# 1. Don't forget that the first argument for specifying a cell in worksheet is a row.  So whereas typically when specifying co-ordinates it is [x,y], with the spreadsheet gem it is [y,x]

require File.expand_path(File.join(File.dirname(__FILE__), '..', 'lib', 'spreadsheet', 'real_coordinates'))

require 'minitest/autorun'

class TC_RealCoordinates < MiniTest::Unit::TestCase

  def setup
    workbook = Spreadsheet.open('test.xls')
    @worksheet = workbook.worksheet(0)
    @worksheet[0,0] = '0-0'
    @worksheet[0,1] = '0-1'
    @worksheet[1,0] = '1-0'
    @worksheet[1,1] = '1-1'
  end

  def test_column_number
    assert_equal 0, @worksheet.column_number('a')
    assert_equal 1, @worksheet.column_number('b')
    assert_equal 2, @worksheet.column_number('c')

    assert_equal 25, @worksheet.column_number('z')
    assert_equal 26, @worksheet.column_number('aa')
    assert_equal 27, @worksheet.column_number('ab')

    assert_equal 51, @worksheet.column_number('az')
    assert_equal 52, @worksheet.column_number('ba')
    assert_equal 53, @worksheet.column_number('bb')

    assert_equal 701, @worksheet.column_number('zz')
    assert_equal 702, @worksheet.column_number('aaa')
    assert_equal 703, @worksheet.column_number('aab')

    assert_equal 727, @worksheet.column_number('aaz')
    assert_equal 728, @worksheet.column_number('aba')
    assert_equal 729, @worksheet.column_number('abb')
  end

  def test_column_letters
    assert_equal 'A', @worksheet.column_letters(0)
    assert_equal 'B', @worksheet.column_letters(1)
    assert_equal 'C', @worksheet.column_letters(2)

    assert_equal 'Z', @worksheet.column_letters(25)
    assert_equal 'AA', @worksheet.column_letters(26)
    assert_equal 'AB', @worksheet.column_letters(27)

    assert_equal 'AZ', @worksheet.column_letters(51)
    assert_equal 'BA', @worksheet.column_letters(52)
    assert_equal 'BB', @worksheet.column_letters(53)

    assert_equal 'ZZ', @worksheet.column_letters(701)
    assert_equal 'AAA', @worksheet.column_letters(702)
    assert_equal 'AAB', @worksheet.column_letters(703)

    assert_equal 'AAZ', @worksheet.column_letters(727)
    assert_equal 'ABA', @worksheet.column_letters(728)
    assert_equal 'ABB', @worksheet.column_letters(729)
  end

  def test_numeric_coordinates
    # Row 1
    assert_equal [0,0], @worksheet.numeric_coordinates('a1')
    assert_equal [0,1], @worksheet.numeric_coordinates('b1')
    assert_equal [0,2], @worksheet.numeric_coordinates('c1')
    assert_equal [0,25], @worksheet.numeric_coordinates('z1')
    assert_equal [0,26], @worksheet.numeric_coordinates('aa1')
    assert_equal [0,27], @worksheet.numeric_coordinates('ab1')
    assert_equal [0,51], @worksheet.numeric_coordinates('az1')
    assert_equal [0,52], @worksheet.numeric_coordinates('ba1')
    assert_equal [0,53], @worksheet.numeric_coordinates('bb1')
    # Row 2
    assert_equal [1,0], @worksheet.numeric_coordinates('a2')
    assert_equal [1,1], @worksheet.numeric_coordinates('b2')
    assert_equal [1,2], @worksheet.numeric_coordinates('c2')
    assert_equal [1,25], @worksheet.numeric_coordinates('z2')
    assert_equal [1,26], @worksheet.numeric_coordinates('aa2')
    assert_equal [1,27], @worksheet.numeric_coordinates('ab2')
    assert_equal [1,51], @worksheet.numeric_coordinates('az2')
    assert_equal [1,52], @worksheet.numeric_coordinates('ba2')
    assert_equal [1,53], @worksheet.numeric_coordinates('bb2')
  end

  def test_method_missing
    assert_equal '0-0', @worksheet.a1
    assert_equal '0-0.changed', @worksheet.a1 = '0-0.changed'
    assert_equal '0-0.changed', @worksheet.a1

    assert_equal '0-1', @worksheet.b1
    assert_equal '0-1.changed', @worksheet.b1 = '0-1.changed'
    assert_equal '0-1.changed', @worksheet.b1

    assert_equal '1-0', @worksheet.a2
    assert_equal '1-0.changed', @worksheet.a2 = '1-0.changed'
    assert_equal '1-0.changed', @worksheet.a2

    assert_equal '1-1', @worksheet.b2
    assert_equal '1-1.changed', @worksheet.b2 = '1-1.changed'
    assert_equal '1-1.changed', @worksheet.b2
  end

end
