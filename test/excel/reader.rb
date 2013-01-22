#!/usr/bin/env ruby
# Excel::TestReader -- Spreadsheet -- 22.01.2013

$: << File.expand_path('../../../lib', File.dirname(__FILE__))

require 'test/unit'

module Spreadsheet
  module Excel
    class TestReader < Test::Unit::TestCase
      def test_empty_file_error_on_setup
        reader = Spreadsheet::Excel::Reader.new
        empty_io = StringIO.new
        assert_raise RuntimeError do
          reader.setup empty_io
        end
      end

      def test_not_empty_file_error_on_setup
        reader = Spreadsheet::Excel::Reader.new
        data = File.expand_path File.join('test', 'data')
        path = File.join data, 'test_empty.xls'
        not_empty_io = File.open(path, "rb")
        assert_nothing_thrown do
          reader.setup not_empty_io
        end
      end
    end
  end
end