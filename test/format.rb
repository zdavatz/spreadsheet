#!/usr/bin/env ruby
# TestFormat -- Spreadsheet -- 06.11.2012 -- mina.git@naguib.ca

$: << File.expand_path('../lib', File.dirname(__FILE__))

require 'test/unit'
require 'spreadsheet'

module Spreadsheet
  class TestFormat < Test::Unit::TestCase
    def setup
      @format = Format.new
    end
    def test_date?
      assert_equal false, @format.date?
      @format.number_format = "hms"
      assert_equal false, @format.date?
      @format.number_format = "Y"
      assert_equal true, @format.date?
      @format.number_format = "YMD"
      assert_equal true, @format.date?
    end
    def test_date_or_time?
      assert_equal false, @format.date_or_time?
      @format.number_format = "hms"
      assert_equal true, @format.date_or_time?
      @format.number_format = "YMD"
      assert_equal true, @format.date_or_time?
      @format.number_format = "hmsYMD"
      assert_equal true, @format.date_or_time?
    end
    def test_datetime?
      assert_equal false, @format.datetime?
      @format.number_format = "H"
      assert_equal false, @format.datetime?
      @format.number_format = "S"
      assert_equal false, @format.datetime?
      @format.number_format = "Y"
      assert_equal false, @format.datetime?
      @format.number_format = "HSYMD"
      assert_equal true, @format.datetime?
    end
    def test_time?
      assert_equal false, @format.time?
      @format.number_format = "YMD"
      assert_equal false, @format.time?
      @format.number_format = "hmsYMD"
      assert_equal true, @format.time?
      @format.number_format = "h"
      assert_equal true, @format.time?
      @format.number_format = "hm"
      assert_equal true, @format.time?
      @format.number_format = "hms"
      assert_equal true, @format.time?
    end
  end
end
