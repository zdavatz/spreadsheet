#!/usr/bin/env ruby
# TestWorkbook -- Spreadsheet -- 24.09.2008 -- hwyss@ywesee.com

$: << File.expand_path('../lib', File.dirname(__FILE__))

require 'test/unit'
require 'spreadsheet'
require 'fileutils'
require 'stringio'

module Spreadsheet
  class TestWorkbook < Test::Unit::TestCase
    def setup
      @io = StringIO.new ''
      @book = Workbook.new
    end
    def test_writer__default_excel
      assert_instance_of Excel::Writer::Workbook, @book.writer(@io)
    end
    def test_sheet_count
        @worksheet1 = Excel::Worksheet.new
        @book.add_worksheet @worksheet1
        assert_equal 1, @book.sheet_count
        @worksheet2 = Excel::Worksheet.new
        @book.add_worksheet @worksheet2
        assert_equal 2, @book.sheet_count
    end
  end
end
