#!/usr/bin/env ruby
# TestWorksheet -- Spreadheet -- 30.09.2008 -- hwyss@ywesee.com

$: << File.expand_path('../lib', File.dirname(__FILE__))

require 'test/unit'
require 'spreadsheet/worksheet'

module Spreadsheet
  class TestWorksheet < Test::Unit::TestCase
    def setup
      @sheet = Worksheet.new
    end
    def test_cell_writer
      assert_nil @sheet[0,0]
      assert_equal 0, @sheet.column_count
      assert_equal 0, @sheet.row_count
      @sheet[0,0] = 'foo'
      assert_equal 'foo', @sheet[0,0]
      assert_equal 1, @sheet.column_count
      assert_equal 1, @sheet.row_count
      @sheet[1,0] = 'bar'
      assert_equal 1, @sheet.column_count
      assert_equal 2, @sheet.row_count
      @sheet[0,1] = 'bar'
      assert_equal 2, @sheet.column_count
      assert_equal 2, @sheet.row_count
      @sheet[1,0] = nil
      assert_equal 2, @sheet.column_count
      assert_equal 1, @sheet.row_count
      @sheet[0,1] = nil
      assert_equal 1, @sheet.column_count
      assert_equal 1, @sheet.row_count
    end
    def test_column_count
      assert_equal 0, @sheet.column_count
      @sheet.replace_row 3, nil, nil, 1, 2, 'foo, bar'
      assert_equal 3, @sheet.column_count
      @sheet.replace_row 8, nil, 'something', 4, 7, nil
      assert_equal 4, @sheet.column_count
      @sheet.replace_row 5, 4, 'something', 4, 7, nil
      assert_equal 5, @sheet.column_count
      @sheet.replace_row 5, nil, 'something', 4, 7, nil
      assert_equal 4, @sheet.column_count
      @sheet.replace_row 3
      assert_equal 3, @sheet.column_count
    end
    def test_row_count
      assert_equal 0, @sheet.row_count
      @sheet.replace_row 3, nil, nil, 1, 2, 'foo, bar'
      assert_equal 1, @sheet.row_count
      @sheet.replace_row 8, nil, 'something', 4, 7, nil
      assert_equal 6, @sheet.row_count
      @sheet.replace_row 5, 4, 'something', 4, 7, nil
      assert_equal 6, @sheet.row_count
      @sheet.replace_row 5, nil, 'something', 4, 7, nil
      assert_equal 6, @sheet.row_count
      @sheet.replace_row 3
      assert_equal 4, @sheet.row_count
    end
  end
end
