#!/usr/bin/env ruby
# TestWorkbook -- Spreadheet -- 24.09.2008 -- hwyss@ywesee.com

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
  end
end
