# test/real_coordinates_test.rb

# Notes:
# 1. Don't forget that the first argument for specifying a cell in worksheet is a row.  So whereas typically when specifying co-ordinates it is [x,y], with the spreadsheet gem it is [y,x].

require File.expand_path(File.join(File.dirname(__FILE__), '..', 'lib', 'real_coordinates'))

require 'minitest/autorun'

class TC_RealCoordinates_Spreadsheet_Worksheet < MiniTest::Unit::TestCase

  def setup
    @workbook = Spreadsheet::Workbook.new
    @worksheet = Spreadsheet::Worksheet.new(workbook: @workbook)
    %w{a b c d e}.each do |column|
      (1..5).each do |row|
        @worksheet.send("#{column}#{row}=", "#{column},#{row}")
      end
    end
  end

  def test_column_number
    assert_equal 0, @worksheet.column_number('a')
    assert_equal 0, @worksheet.column_number('A')
    assert_equal 1, @worksheet.column_number('b')
    assert_equal 1, @worksheet.column_number('B')
    assert_equal 2, @worksheet.column_number('c')
    assert_equal 2, @worksheet.column_number('C')
    assert_equal 24, @worksheet.column_number('y')
    assert_equal 24, @worksheet.column_number('Y')
    assert_equal 25, @worksheet.column_number('z')
    assert_equal 25, @worksheet.column_number('Z')

    assert_equal 26, @worksheet.column_number('aa')
    assert_equal 26, @worksheet.column_number('AA')
    assert_equal 27, @worksheet.column_number('ab')
    assert_equal 27, @worksheet.column_number('AB')
    assert_equal 28, @worksheet.column_number('ac')
    assert_equal 28, @worksheet.column_number('AC')
    assert_equal 50, @worksheet.column_number('ay')
    assert_equal 50, @worksheet.column_number('AY')
    assert_equal 51, @worksheet.column_number('az')
    assert_equal 51, @worksheet.column_number('AZ')

    assert_equal 52, @worksheet.column_number('ba')
    assert_equal 52, @worksheet.column_number('BA')
    assert_equal 53, @worksheet.column_number('bb')
    assert_equal 53, @worksheet.column_number('BB')
    assert_equal 54, @worksheet.column_number('bc')
    assert_equal 54, @worksheet.column_number('BC')
    assert_equal 76, @worksheet.column_number('by')
    assert_equal 76, @worksheet.column_number('BY')
    assert_equal 77, @worksheet.column_number('bz')
    assert_equal 77, @worksheet.column_number('BZ')

    assert_equal 78, @worksheet.column_number('ca')
    assert_equal 78, @worksheet.column_number('CA')
    assert_equal 79, @worksheet.column_number('cb')
    assert_equal 79, @worksheet.column_number('CB')
    assert_equal 80, @worksheet.column_number('cc')
    assert_equal 80, @worksheet.column_number('CC')
    assert_equal 102, @worksheet.column_number('cy')
    assert_equal 102, @worksheet.column_number('CY')
    assert_equal 103, @worksheet.column_number('cz')
    assert_equal 103, @worksheet.column_number('CZ')

    assert_equal 650, @worksheet.column_number('ya')
    assert_equal 650, @worksheet.column_number('YA')
    assert_equal 651, @worksheet.column_number('yb')
    assert_equal 651, @worksheet.column_number('YB')
    assert_equal 652, @worksheet.column_number('yc')
    assert_equal 652, @worksheet.column_number('YC')
    assert_equal 674, @worksheet.column_number('yy')
    assert_equal 674, @worksheet.column_number('YY')
    assert_equal 675, @worksheet.column_number('yz')
    assert_equal 675, @worksheet.column_number('YZ')

    assert_equal 676, @worksheet.column_number('za')
    assert_equal 676, @worksheet.column_number('ZA')
    assert_equal 677, @worksheet.column_number('zb')
    assert_equal 677, @worksheet.column_number('ZB')
    assert_equal 678, @worksheet.column_number('zc')
    assert_equal 678, @worksheet.column_number('ZC')
    assert_equal 700, @worksheet.column_number('zy')
    assert_equal 700, @worksheet.column_number('ZY')
    assert_equal 701, @worksheet.column_number('zz')
    assert_equal 701, @worksheet.column_number('ZZ')

    assert_equal 702, @worksheet.column_number('aaa')
    assert_equal 702, @worksheet.column_number('AAA')
    assert_equal 703, @worksheet.column_number('aab')
    assert_equal 703, @worksheet.column_number('AAB')
    assert_equal 704, @worksheet.column_number('aac')
    assert_equal 704, @worksheet.column_number('AAC')
    assert_equal 726, @worksheet.column_number('aay')
    assert_equal 726, @worksheet.column_number('AAY')
    assert_equal 727, @worksheet.column_number('aaz')
    assert_equal 727, @worksheet.column_number('AAZ')

    assert_equal 728, @worksheet.column_number('aba')
    assert_equal 728, @worksheet.column_number('ABA')
    assert_equal 729, @worksheet.column_number('abb')
    assert_equal 729, @worksheet.column_number('ABB')
    assert_equal 730, @worksheet.column_number('abc')
    assert_equal 730, @worksheet.column_number('ABC')
    assert_equal 752, @worksheet.column_number('aby')
    assert_equal 752, @worksheet.column_number('ABY')
    assert_equal 753, @worksheet.column_number('abz')
    assert_equal 753, @worksheet.column_number('ABZ')
  end

  def test_place_values
    assert_equal [0], @worksheet.place_values(0)
    assert_equal [1], @worksheet.place_values(1)
    assert_equal [2], @worksheet.place_values(2)
    assert_equal [24], @worksheet.place_values(24)
    assert_equal [25], @worksheet.place_values(25)

    assert_equal [0,0], @worksheet.place_values(26)
    assert_equal [0,1], @worksheet.place_values(27)
    assert_equal [0,2], @worksheet.place_values(28)
    assert_equal [0,24], @worksheet.place_values(50)
    assert_equal [0,25], @worksheet.place_values(51)

    assert_equal [1,0], @worksheet.place_values(52)
    assert_equal [1,1], @worksheet.place_values(53)
    assert_equal [1,2], @worksheet.place_values(54)
    assert_equal [1,24], @worksheet.place_values(76)
    assert_equal [1,25], @worksheet.place_values(77)

    assert_equal [2,0], @worksheet.place_values(78)
    assert_equal [2,1], @worksheet.place_values(79)
    assert_equal [2,2], @worksheet.place_values(80)
    assert_equal [2,24], @worksheet.place_values(102)
    assert_equal [2,25], @worksheet.place_values(103)

    assert_equal [24,0], @worksheet.place_values(650)
    assert_equal [24,1], @worksheet.place_values(651)
    assert_equal [24,2], @worksheet.place_values(652)
    assert_equal [24,24], @worksheet.place_values(674)
    assert_equal [24,25], @worksheet.place_values(675)

    assert_equal [25,0], @worksheet.place_values(676)
    assert_equal [25,1], @worksheet.place_values(677)
    assert_equal [25,2], @worksheet.place_values(678)
    assert_equal [25,24], @worksheet.place_values(700)
    assert_equal [25,25], @worksheet.place_values(701)

    assert_equal [0,0,0], @worksheet.place_values(702)
    assert_equal [0,0,1], @worksheet.place_values(703)
    assert_equal [0,0,2], @worksheet.place_values(704)
    assert_equal [0,0,24], @worksheet.place_values(726)
    assert_equal [0,0,25], @worksheet.place_values(727)

    assert_equal [0,1,0], @worksheet.place_values(728)
    assert_equal [0,1,1], @worksheet.place_values(729)
    assert_equal [0,1,2], @worksheet.place_values(730)
    assert_equal [0,1,24], @worksheet.place_values(752)
    assert_equal [0,1,25], @worksheet.place_values(753)
  end

  def test_column_letters
    assert_equal 'A', @worksheet.column_letters(0)
    assert_equal 'B', @worksheet.column_letters(1)
    assert_equal 'C', @worksheet.column_letters(2)
    assert_equal 'Y', @worksheet.column_letters(24)
    assert_equal 'Z', @worksheet.column_letters(25)

    assert_equal 'AA', @worksheet.column_letters(26)
    assert_equal 'AB', @worksheet.column_letters(27)
    assert_equal 'AC', @worksheet.column_letters(28)
    assert_equal 'AY', @worksheet.column_letters(50)
    assert_equal 'AZ', @worksheet.column_letters(51)

    assert_equal 'BA', @worksheet.column_letters(52)
    assert_equal 'BB', @worksheet.column_letters(53)
    assert_equal 'BC', @worksheet.column_letters(54)
    assert_equal 'BY', @worksheet.column_letters(76)
    assert_equal 'BZ', @worksheet.column_letters(77)

    assert_equal 'CA', @worksheet.column_letters(78)
    assert_equal 'CB', @worksheet.column_letters(79)
    assert_equal 'CC', @worksheet.column_letters(80)
    assert_equal 'CY', @worksheet.column_letters(102)
    assert_equal 'CZ', @worksheet.column_letters(103)

    assert_equal 'YA', @worksheet.column_letters(650)
    assert_equal 'YB', @worksheet.column_letters(651)
    assert_equal 'YC', @worksheet.column_letters(652)
    assert_equal 'YY', @worksheet.column_letters(674)
    assert_equal 'YZ', @worksheet.column_letters(675)

    assert_equal 'ZA', @worksheet.column_letters(676)
    assert_equal 'ZB', @worksheet.column_letters(677)
    assert_equal 'ZC', @worksheet.column_letters(678)
    assert_equal 'ZY', @worksheet.column_letters(700)
    assert_equal 'ZZ', @worksheet.column_letters(701)

    assert_equal 'AAA', @worksheet.column_letters(702)
    assert_equal 'AAB', @worksheet.column_letters(703)
    assert_equal 'AAC', @worksheet.column_letters(704)
    assert_equal 'AAY', @worksheet.column_letters(726)
    assert_equal 'AAZ', @worksheet.column_letters(727)

    assert_equal 'ABA', @worksheet.column_letters(728)
    assert_equal 'ABB', @worksheet.column_letters(729)
    assert_equal 'ABC', @worksheet.column_letters(730)
    assert_equal 'ABY', @worksheet.column_letters(752)
    assert_equal 'ABZ', @worksheet.column_letters(753)
  end

  def test_numeric_coordinates

    # Row 1

    assert_equal [0,0], @worksheet.numeric_coordinates('a1')
    assert_equal [0,0], @worksheet.numeric_coordinates('A1')
    assert_equal [0,1], @worksheet.numeric_coordinates('b1')
    assert_equal [0,1], @worksheet.numeric_coordinates('B1')
    assert_equal [0,2], @worksheet.numeric_coordinates('c1')
    assert_equal [0,2], @worksheet.numeric_coordinates('C1')
    assert_equal [0,24], @worksheet.numeric_coordinates('y1')
    assert_equal [0,24], @worksheet.numeric_coordinates('Y1')
    assert_equal [0,25], @worksheet.numeric_coordinates('z1')
    assert_equal [0,25], @worksheet.numeric_coordinates('Z1')

    assert_equal [0,26], @worksheet.numeric_coordinates('aa1')
    assert_equal [0,26], @worksheet.numeric_coordinates('AA1')
    assert_equal [0,27], @worksheet.numeric_coordinates('ab1')
    assert_equal [0,27], @worksheet.numeric_coordinates('AB1')
    assert_equal [0,28], @worksheet.numeric_coordinates('ac1')
    assert_equal [0,28], @worksheet.numeric_coordinates('AC1')
    assert_equal [0,50], @worksheet.numeric_coordinates('ay1')
    assert_equal [0,50], @worksheet.numeric_coordinates('AY1')
    assert_equal [0,51], @worksheet.numeric_coordinates('az1')
    assert_equal [0,51], @worksheet.numeric_coordinates('AZ1')

    assert_equal [0,52], @worksheet.numeric_coordinates('ba1')
    assert_equal [0,52], @worksheet.numeric_coordinates('BA1')
    assert_equal [0,53], @worksheet.numeric_coordinates('bb1')
    assert_equal [0,53], @worksheet.numeric_coordinates('BB1')
    assert_equal [0,54], @worksheet.numeric_coordinates('bc1')
    assert_equal [0,54], @worksheet.numeric_coordinates('BC1')
    assert_equal [0,76], @worksheet.numeric_coordinates('by1')
    assert_equal [0,76], @worksheet.numeric_coordinates('BY1')
    assert_equal [0,77], @worksheet.numeric_coordinates('bz1')
    assert_equal [0,77], @worksheet.numeric_coordinates('BZ1')

    assert_equal [0,78], @worksheet.numeric_coordinates('ca1')
    assert_equal [0,78], @worksheet.numeric_coordinates('CA1')
    assert_equal [0,79], @worksheet.numeric_coordinates('cb1')
    assert_equal [0,79], @worksheet.numeric_coordinates('CB1')
    assert_equal [0,80], @worksheet.numeric_coordinates('cc1')
    assert_equal [0,80], @worksheet.numeric_coordinates('CC1')
    assert_equal [0,102], @worksheet.numeric_coordinates('cy1')
    assert_equal [0,102], @worksheet.numeric_coordinates('CY1')
    assert_equal [0,103], @worksheet.numeric_coordinates('cz1')
    assert_equal [0,103], @worksheet.numeric_coordinates('CZ1')

    assert_equal [0,650], @worksheet.numeric_coordinates('ya1')
    assert_equal [0,650], @worksheet.numeric_coordinates('YA1')
    assert_equal [0,651], @worksheet.numeric_coordinates('yb1')
    assert_equal [0,651], @worksheet.numeric_coordinates('YB1')
    assert_equal [0,652], @worksheet.numeric_coordinates('yc1')
    assert_equal [0,652], @worksheet.numeric_coordinates('YC1')
    assert_equal [0,674], @worksheet.numeric_coordinates('yy1')
    assert_equal [0,674], @worksheet.numeric_coordinates('YY1')
    assert_equal [0,675], @worksheet.numeric_coordinates('yz1')
    assert_equal [0,675], @worksheet.numeric_coordinates('YZ1')

    assert_equal [0,676], @worksheet.numeric_coordinates('za1')
    assert_equal [0,676], @worksheet.numeric_coordinates('ZA1')
    assert_equal [0,677], @worksheet.numeric_coordinates('zb1')
    assert_equal [0,677], @worksheet.numeric_coordinates('ZB1')
    assert_equal [0,678], @worksheet.numeric_coordinates('zc1')
    assert_equal [0,678], @worksheet.numeric_coordinates('ZC1')
    assert_equal [0,700], @worksheet.numeric_coordinates('zy1')
    assert_equal [0,700], @worksheet.numeric_coordinates('ZY1')
    assert_equal [0,701], @worksheet.numeric_coordinates('zz1')
    assert_equal [0,701], @worksheet.numeric_coordinates('ZZ1')

    assert_equal [0,702], @worksheet.numeric_coordinates('aaa1')
    assert_equal [0,702], @worksheet.numeric_coordinates('AAA1')
    assert_equal [0,703], @worksheet.numeric_coordinates('aab1')
    assert_equal [0,703], @worksheet.numeric_coordinates('AAB1')
    assert_equal [0,704], @worksheet.numeric_coordinates('aac1')
    assert_equal [0,704], @worksheet.numeric_coordinates('AAC1')
    assert_equal [0,726], @worksheet.numeric_coordinates('aay1')
    assert_equal [0,726], @worksheet.numeric_coordinates('AAY1')
    assert_equal [0,727], @worksheet.numeric_coordinates('aaz1')
    assert_equal [0,727], @worksheet.numeric_coordinates('AAZ1')

    assert_equal [0,728], @worksheet.numeric_coordinates('aba1')
    assert_equal [0,728], @worksheet.numeric_coordinates('ABA1')
    assert_equal [0,729], @worksheet.numeric_coordinates('abb1')
    assert_equal [0,729], @worksheet.numeric_coordinates('ABB1')
    assert_equal [0,730], @worksheet.numeric_coordinates('abc1')
    assert_equal [0,730], @worksheet.numeric_coordinates('ABC1')
    assert_equal [0,752], @worksheet.numeric_coordinates('aby1')
    assert_equal [0,752], @worksheet.numeric_coordinates('ABY1')
    assert_equal [0,753], @worksheet.numeric_coordinates('abz1')
    assert_equal [0,753], @worksheet.numeric_coordinates('ABZ1')

    # Row 2

    assert_equal [1,0], @worksheet.numeric_coordinates('a2')
    assert_equal [1,0], @worksheet.numeric_coordinates('A2')
    assert_equal [1,1], @worksheet.numeric_coordinates('b2')
    assert_equal [1,1], @worksheet.numeric_coordinates('B2')
    assert_equal [1,2], @worksheet.numeric_coordinates('c2')
    assert_equal [1,2], @worksheet.numeric_coordinates('C2')
    assert_equal [1,24], @worksheet.numeric_coordinates('y2')
    assert_equal [1,24], @worksheet.numeric_coordinates('Y2')
    assert_equal [1,25], @worksheet.numeric_coordinates('z2')
    assert_equal [1,25], @worksheet.numeric_coordinates('Z2')

    assert_equal [1,26], @worksheet.numeric_coordinates('aa2')
    assert_equal [1,26], @worksheet.numeric_coordinates('AA2')
    assert_equal [1,27], @worksheet.numeric_coordinates('ab2')
    assert_equal [1,27], @worksheet.numeric_coordinates('AB2')
    assert_equal [1,28], @worksheet.numeric_coordinates('ac2')
    assert_equal [1,28], @worksheet.numeric_coordinates('AC2')
    assert_equal [1,50], @worksheet.numeric_coordinates('ay2')
    assert_equal [1,50], @worksheet.numeric_coordinates('AY2')
    assert_equal [1,51], @worksheet.numeric_coordinates('az2')
    assert_equal [1,51], @worksheet.numeric_coordinates('AZ2')

    assert_equal [1,52], @worksheet.numeric_coordinates('ba2')
    assert_equal [1,52], @worksheet.numeric_coordinates('BA2')
    assert_equal [1,53], @worksheet.numeric_coordinates('bb2')
    assert_equal [1,53], @worksheet.numeric_coordinates('BB2')
    assert_equal [1,54], @worksheet.numeric_coordinates('bc2')
    assert_equal [1,54], @worksheet.numeric_coordinates('BC2')
    assert_equal [1,76], @worksheet.numeric_coordinates('by2')
    assert_equal [1,76], @worksheet.numeric_coordinates('BY2')
    assert_equal [1,77], @worksheet.numeric_coordinates('bz2')
    assert_equal [1,77], @worksheet.numeric_coordinates('BZ2')

    assert_equal [1,78], @worksheet.numeric_coordinates('ca2')
    assert_equal [1,78], @worksheet.numeric_coordinates('CA2')
    assert_equal [1,79], @worksheet.numeric_coordinates('cb2')
    assert_equal [1,79], @worksheet.numeric_coordinates('CB2')
    assert_equal [1,80], @worksheet.numeric_coordinates('cc2')
    assert_equal [1,80], @worksheet.numeric_coordinates('CC2')
    assert_equal [1,102], @worksheet.numeric_coordinates('cy2')
    assert_equal [1,102], @worksheet.numeric_coordinates('CY2')
    assert_equal [1,103], @worksheet.numeric_coordinates('cz2')
    assert_equal [1,103], @worksheet.numeric_coordinates('CZ2')

    assert_equal [1,650], @worksheet.numeric_coordinates('ya2')
    assert_equal [1,650], @worksheet.numeric_coordinates('YA2')
    assert_equal [1,651], @worksheet.numeric_coordinates('yb2')
    assert_equal [1,651], @worksheet.numeric_coordinates('YB2')
    assert_equal [1,652], @worksheet.numeric_coordinates('yc2')
    assert_equal [1,652], @worksheet.numeric_coordinates('YC2')
    assert_equal [1,674], @worksheet.numeric_coordinates('yy2')
    assert_equal [1,674], @worksheet.numeric_coordinates('YY2')
    assert_equal [1,675], @worksheet.numeric_coordinates('yz2')
    assert_equal [1,675], @worksheet.numeric_coordinates('YZ2')

    assert_equal [1,676], @worksheet.numeric_coordinates('za2')
    assert_equal [1,676], @worksheet.numeric_coordinates('ZA2')
    assert_equal [1,677], @worksheet.numeric_coordinates('zb2')
    assert_equal [1,677], @worksheet.numeric_coordinates('ZB2')
    assert_equal [1,678], @worksheet.numeric_coordinates('zc2')
    assert_equal [1,678], @worksheet.numeric_coordinates('ZC2')
    assert_equal [1,700], @worksheet.numeric_coordinates('zy2')
    assert_equal [1,700], @worksheet.numeric_coordinates('ZY2')
    assert_equal [1,701], @worksheet.numeric_coordinates('zz2')
    assert_equal [1,701], @worksheet.numeric_coordinates('ZZ2')

    assert_equal [1,702], @worksheet.numeric_coordinates('aaa2')
    assert_equal [1,702], @worksheet.numeric_coordinates('AAA2')
    assert_equal [1,703], @worksheet.numeric_coordinates('aab2')
    assert_equal [1,703], @worksheet.numeric_coordinates('AAB2')
    assert_equal [1,704], @worksheet.numeric_coordinates('aac2')
    assert_equal [1,704], @worksheet.numeric_coordinates('AAC2')
    assert_equal [1,726], @worksheet.numeric_coordinates('aay2')
    assert_equal [1,726], @worksheet.numeric_coordinates('AAY2')
    assert_equal [1,727], @worksheet.numeric_coordinates('aaz2')
    assert_equal [1,727], @worksheet.numeric_coordinates('AAZ2')

    assert_equal [1,728], @worksheet.numeric_coordinates('aba2')
    assert_equal [1,728], @worksheet.numeric_coordinates('ABA2')
    assert_equal [1,729], @worksheet.numeric_coordinates('abb2')
    assert_equal [1,729], @worksheet.numeric_coordinates('ABB2')
    assert_equal [1,730], @worksheet.numeric_coordinates('abc2')
    assert_equal [1,730], @worksheet.numeric_coordinates('ABC2')
    assert_equal [1,752], @worksheet.numeric_coordinates('aby2')
    assert_equal [1,752], @worksheet.numeric_coordinates('ABY2')
    assert_equal [1,753], @worksheet.numeric_coordinates('abz2')
    assert_equal [1,753], @worksheet.numeric_coordinates('ABZ2')
  end

  def test_real_coordinates

    # Row 1

    assert_equal 'A1', @worksheet.real_coordinates(0,0)
    assert_equal 'A1', @worksheet.real_coordinates([0,0])
    assert_equal 'B1', @worksheet.real_coordinates(0,1)
    assert_equal 'B1', @worksheet.real_coordinates([0,1])
    assert_equal 'C1', @worksheet.real_coordinates(0,2)
    assert_equal 'C1', @worksheet.real_coordinates([0,2])
    assert_equal 'Z1', @worksheet.real_coordinates(0,25)
    assert_equal 'Z1', @worksheet.real_coordinates([0,25])

    assert_equal 'AA1', @worksheet.real_coordinates(0,26)
    assert_equal 'AA1', @worksheet.real_coordinates([0,26])
    assert_equal 'AB1', @worksheet.real_coordinates(0,27)
    assert_equal 'AB1', @worksheet.real_coordinates([0,27])
    assert_equal 'AC1', @worksheet.real_coordinates(0,28)
    assert_equal 'AC1', @worksheet.real_coordinates([0,28])
    assert_equal 'AZ1', @worksheet.real_coordinates(0,51)
    assert_equal 'AZ1', @worksheet.real_coordinates([0,51])

    assert_equal 'BA1', @worksheet.real_coordinates(0,52)
    assert_equal 'BA1', @worksheet.real_coordinates([0,52])
    assert_equal 'BB1', @worksheet.real_coordinates(0,53)
    assert_equal 'BB1', @worksheet.real_coordinates([0,53])
    assert_equal 'BC1', @worksheet.real_coordinates(0,54)
    assert_equal 'BC1', @worksheet.real_coordinates([0,54])
    assert_equal 'BZ1', @worksheet.real_coordinates(0,77)
    assert_equal 'BZ1', @worksheet.real_coordinates([0,77])

    assert_equal 'CA1', @worksheet.real_coordinates(0,78)
    assert_equal 'CA1', @worksheet.real_coordinates([0,78])
    assert_equal 'CB1', @worksheet.real_coordinates(0,79)
    assert_equal 'CB1', @worksheet.real_coordinates([0,79])
    assert_equal 'CC1', @worksheet.real_coordinates(0,80)
    assert_equal 'CC1', @worksheet.real_coordinates([0,80])
    assert_equal 'CY1', @worksheet.real_coordinates(0,102)
    assert_equal 'CY1', @worksheet.real_coordinates([0,102])
    assert_equal 'CZ1', @worksheet.real_coordinates(0,103)
    assert_equal 'CZ1', @worksheet.real_coordinates([0,103])

    assert_equal 'YA1', @worksheet.real_coordinates(0,650)
    assert_equal 'YA1', @worksheet.real_coordinates([0,650])
    assert_equal 'YB1', @worksheet.real_coordinates(0,651)
    assert_equal 'YB1', @worksheet.real_coordinates([0,651])
    assert_equal 'YC1', @worksheet.real_coordinates(0,652)
    assert_equal 'YC1', @worksheet.real_coordinates([0,652])
    assert_equal 'YY1', @worksheet.real_coordinates(0,674)
    assert_equal 'YY1', @worksheet.real_coordinates([0,674])
    assert_equal 'YZ1', @worksheet.real_coordinates(0,675)
    assert_equal 'YZ1', @worksheet.real_coordinates([0,675])

    assert_equal 'ZA1', @worksheet.real_coordinates(0,676)
    assert_equal 'ZA1', @worksheet.real_coordinates([0,676])
    assert_equal 'ZB1', @worksheet.real_coordinates(0,677)
    assert_equal 'ZB1', @worksheet.real_coordinates([0,677])
    assert_equal 'ZC1', @worksheet.real_coordinates(0,678)
    assert_equal 'ZC1', @worksheet.real_coordinates([0,678])
    assert_equal 'ZY1', @worksheet.real_coordinates(0,700)
    assert_equal 'ZY1', @worksheet.real_coordinates([0,700])
    assert_equal 'ZZ1', @worksheet.real_coordinates(0,701)
    assert_equal 'ZZ1', @worksheet.real_coordinates([0,701])

    assert_equal 'AAA1', @worksheet.real_coordinates(0,702)
    assert_equal 'AAA1', @worksheet.real_coordinates([0,702])
    assert_equal 'AAB1', @worksheet.real_coordinates(0,703)
    assert_equal 'AAB1', @worksheet.real_coordinates([0,703])
    assert_equal 'AAC1', @worksheet.real_coordinates(0,704)
    assert_equal 'AAC1', @worksheet.real_coordinates([0,704])
    assert_equal 'AAY1', @worksheet.real_coordinates(0,726)
    assert_equal 'AAY1', @worksheet.real_coordinates([0,726])
    assert_equal 'AAZ1', @worksheet.real_coordinates(0,727)
    assert_equal 'AAZ1', @worksheet.real_coordinates([0,727])

    assert_equal 'ABA1', @worksheet.real_coordinates(0,728)
    assert_equal 'ABA1', @worksheet.real_coordinates([0,728])
    assert_equal 'ABB1', @worksheet.real_coordinates(0,729)
    assert_equal 'ABB1', @worksheet.real_coordinates([0,729])
    assert_equal 'ABC1', @worksheet.real_coordinates(0,730)
    assert_equal 'ABC1', @worksheet.real_coordinates([0,730])
    assert_equal 'ABY1', @worksheet.real_coordinates(0,752)
    assert_equal 'ABY1', @worksheet.real_coordinates([0,752])
    assert_equal 'ABZ1', @worksheet.real_coordinates(0,753)
    assert_equal 'ABZ1', @worksheet.real_coordinates([0,753])

    # Row 2

    assert_equal 'A2', @worksheet.real_coordinates(1,0)
    assert_equal 'A2', @worksheet.real_coordinates([1,0])
    assert_equal 'B2', @worksheet.real_coordinates(1,1)
    assert_equal 'B2', @worksheet.real_coordinates([1,1])
    assert_equal 'C2', @worksheet.real_coordinates(1,2)
    assert_equal 'C2', @worksheet.real_coordinates([1,2])
    assert_equal 'Z2', @worksheet.real_coordinates(1,25)
    assert_equal 'Z2', @worksheet.real_coordinates([1,25])

    assert_equal 'AA2', @worksheet.real_coordinates(1,26)
    assert_equal 'AA2', @worksheet.real_coordinates([1,26])
    assert_equal 'AB2', @worksheet.real_coordinates(1,27)
    assert_equal 'AB2', @worksheet.real_coordinates([1,27])
    assert_equal 'AC2', @worksheet.real_coordinates(1,28)
    assert_equal 'AC2', @worksheet.real_coordinates([1,28])
    assert_equal 'AZ2', @worksheet.real_coordinates(1,51)
    assert_equal 'AZ2', @worksheet.real_coordinates([1,51])

    assert_equal 'BA2', @worksheet.real_coordinates(1,52)
    assert_equal 'BA2', @worksheet.real_coordinates([1,52])
    assert_equal 'BB2', @worksheet.real_coordinates(1,53)
    assert_equal 'BB2', @worksheet.real_coordinates([1,53])
    assert_equal 'BC2', @worksheet.real_coordinates(1,54)
    assert_equal 'BC2', @worksheet.real_coordinates([1,54])
    assert_equal 'BZ2', @worksheet.real_coordinates(1,77)
    assert_equal 'BZ2', @worksheet.real_coordinates([1,77])

    assert_equal 'CA2', @worksheet.real_coordinates(1,78)
    assert_equal 'CA2', @worksheet.real_coordinates([1,78])
    assert_equal 'CB2', @worksheet.real_coordinates(1,79)
    assert_equal 'CB2', @worksheet.real_coordinates([1,79])
    assert_equal 'CC2', @worksheet.real_coordinates(1,80)
    assert_equal 'CC2', @worksheet.real_coordinates([1,80])
    assert_equal 'CY2', @worksheet.real_coordinates(1,102)
    assert_equal 'CY2', @worksheet.real_coordinates([1,102])
    assert_equal 'CZ2', @worksheet.real_coordinates(1,103)
    assert_equal 'CZ2', @worksheet.real_coordinates([1,103])

    assert_equal 'YA2', @worksheet.real_coordinates(1,650)
    assert_equal 'YA2', @worksheet.real_coordinates([1,650])
    assert_equal 'YB2', @worksheet.real_coordinates(1,651)
    assert_equal 'YB2', @worksheet.real_coordinates([1,651])
    assert_equal 'YC2', @worksheet.real_coordinates(1,652)
    assert_equal 'YC2', @worksheet.real_coordinates([1,652])
    assert_equal 'YY2', @worksheet.real_coordinates(1,674)
    assert_equal 'YY2', @worksheet.real_coordinates([1,674])
    assert_equal 'YZ2', @worksheet.real_coordinates(1,675)
    assert_equal 'YZ2', @worksheet.real_coordinates([1,675])

    assert_equal 'ZA2', @worksheet.real_coordinates(1,676)
    assert_equal 'ZA2', @worksheet.real_coordinates([1,676])
    assert_equal 'ZB2', @worksheet.real_coordinates(1,677)
    assert_equal 'ZB2', @worksheet.real_coordinates([1,677])
    assert_equal 'ZC2', @worksheet.real_coordinates(1,678)
    assert_equal 'ZC2', @worksheet.real_coordinates([1,678])
    assert_equal 'ZY2', @worksheet.real_coordinates(1,700)
    assert_equal 'ZY2', @worksheet.real_coordinates([1,700])
    assert_equal 'ZZ2', @worksheet.real_coordinates(1,701)
    assert_equal 'ZZ2', @worksheet.real_coordinates([1,701])

    assert_equal 'AAA2', @worksheet.real_coordinates(1,702)
    assert_equal 'AAA2', @worksheet.real_coordinates([1,702])
    assert_equal 'AAB2', @worksheet.real_coordinates(1,703)
    assert_equal 'AAB2', @worksheet.real_coordinates([1,703])
    assert_equal 'AAC2', @worksheet.real_coordinates(1,704)
    assert_equal 'AAC2', @worksheet.real_coordinates([1,704])
    assert_equal 'AAY2', @worksheet.real_coordinates(1,726)
    assert_equal 'AAY2', @worksheet.real_coordinates([1,726])
    assert_equal 'AAZ2', @worksheet.real_coordinates(1,727)
    assert_equal 'AAZ2', @worksheet.real_coordinates([1,727])

    assert_equal 'ABA2', @worksheet.real_coordinates(1,728)
    assert_equal 'ABA2', @worksheet.real_coordinates([1,728])
    assert_equal 'ABB2', @worksheet.real_coordinates(1,729)
    assert_equal 'ABB2', @worksheet.real_coordinates([1,729])
    assert_equal 'ABC2', @worksheet.real_coordinates(1,730)
    assert_equal 'ABC2', @worksheet.real_coordinates([1,730])
    assert_equal 'ABY2', @worksheet.real_coordinates(1,752)
    assert_equal 'ABY2', @worksheet.real_coordinates([1,752])
    assert_equal 'ABZ2', @worksheet.real_coordinates(1,753)
    assert_equal 'ABZ2', @worksheet.real_coordinates([1,753])

  end

  def test_method_missing

    # setup() values

    assert_equal 'a,1', @worksheet.a1
    assert_equal 'a,1', @worksheet[0,0]
    @worksheet.a1 = 'a,1.changed'
    assert_equal 'a,1.changed', @worksheet.a1
    assert_equal 'a,1.changed', @worksheet[0,0]

    assert_equal 'b,1', @worksheet.b1
    assert_equal 'b,1', @worksheet[0,1]
    @worksheet.b1 = 'b,1.changed'
    assert_equal 'b,1.changed', @worksheet.b1
    assert_equal 'b,1.changed', @worksheet[0,1]

    assert_equal 'a,2', @worksheet.a2
    assert_equal 'a,2', @worksheet[1,0]
    @worksheet.a2 = 'a,2.changed'
    assert_equal 'a,2.changed', @worksheet[1,0]

    assert_equal 'b,2', @worksheet.b2
    assert_equal 'b,2', @worksheet[1,1]
    @worksheet.b2 = 'b,2.changed'
    assert_equal 'b,2.changed', @worksheet.b2
    assert_equal 'b,2.changed', @worksheet[1,1]

    assert_equal 'c,1', @worksheet.c1
    assert_equal 'c,1', @worksheet[0,2]
    @worksheet.c1 = 'c,1.changed'
    assert_equal 'c,1.changed', @worksheet.c1
    assert_equal 'c,1.changed', @worksheet[0,2]

    assert_equal 'c,2', @worksheet.c2
    assert_equal 'c,2', @worksheet[1,2]
    @worksheet.c2 = 'c,2.changed'
    assert_equal 'c,2.changed', @worksheet.c2
    assert_equal 'c,2.changed', @worksheet[1,2]

    # non-setup() values

    @worksheet.aa1 = 'aa,1.changed'
    assert_equal 'aa,1.changed', @worksheet.aa1
    assert_equal 'aa,1.changed', @worksheet[0,26]

    @worksheet.aa2 = 'aa,2.changed'
    assert_equal 'aa,2.changed', @worksheet.aa2
    assert_equal 'aa,2.changed', @worksheet[1,26]

    @worksheet.aa3 = 'aa,3.changed'
    assert_equal 'aa,3.changed', @worksheet.aa3
    assert_equal 'aa,3.changed', @worksheet[2,26]

    @worksheet.ab1 = 'ab,1.changed'
    assert_equal 'ab,1.changed', @worksheet.ab1
    assert_equal 'ab,1.changed', @worksheet[0,27]

    @worksheet.ab2 = 'ab,2.changed'
    assert_equal 'ab,2.changed', @worksheet.ab2
    assert_equal 'ab,2.changed', @worksheet[1,27]

    @worksheet.ab3 = 'ab,3.changed'
    assert_equal 'ab,3.changed', @worksheet.ab3
    assert_equal 'ab,3.changed', @worksheet[2,27]
  end

end
