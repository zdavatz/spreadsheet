#!/usr/bin/env ruby
# Excel::TestRow -- Spreadsheet -- 12.10.2008 -- hwyss@ywesee.com

$: << File.expand_path('../../lib', File.dirname(__FILE__))

require 'test/unit'
require 'spreadsheet'

module Spreadsheet
  module Excel
class TestRow < Test::Unit::TestCase
  def setup
    @workbook = Excel::Workbook.new
    @worksheet = Excel::Worksheet.new
    @workbook.add_worksheet @worksheet
  end
  def test_date
    row = Row.new @worksheet, 0, [nil, 27627.6789]
    assert_equal Date.new(1975,8,21), row.date(1)
  end
  def test_datetime
    row = Row.new @worksheet, 0, [nil, 27627.765]
    d1 = DateTime.new(1975,8,21) + 0.765
    d2 = row.datetime 1
    assert_equal d1, d2
  end
end
  end
end
