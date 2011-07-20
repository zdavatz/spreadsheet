#!/usr/bin/env ruby
# Excel::Writer::TestWorkbook -- Spreadsheet -- 20.07.2011 -- vanderhoorn@gmail.com

$: << File.expand_path('../../../lib', File.dirname(__FILE__))

require 'test/unit'
require 'spreadsheet'

module Spreadsheet
  module Excel
    module Writer
      class TestWorkbook < Test::Unit::TestCase
        def test_sanitize_worksheets
          book = Spreadsheet::Excel::Workbook.new
          assert_instance_of Excel::Workbook, book
          assert_equal book.worksheets.size, 0
          workbook_writer = Excel::Writer::Workbook.new book
          assert_nothing_raised { workbook_writer.sanitize_worksheets book.worksheets }
        end
      end
    end
  end
end
