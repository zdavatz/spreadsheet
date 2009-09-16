#!/usr/bin/env ruby
# encoding: utf-8
# TestIntegration -- Spreadheet -- 08.10.2007 -- hwyss@ywesee.com

$: << File.expand_path('../lib', File.dirname(__FILE__))

require 'test/unit'
require 'spreadsheet'
require 'fileutils'

module Spreadsheet
  class TestIntegration < Test::Unit::TestCase
    if RUBY_VERSION >= '1.9'
      class IconvStub
        def initialize to, from
          @to, @from = to, from
        end
        def iconv str
          dp = str.dup
          dp.force_encoding @from
          dp.encode @to
        end
      end
      @@iconv = IconvStub.new('UTF-16LE', 'UTF-8')
      @@bytesize = :bytesize
    else
      @@iconv = Iconv.new('UTF-16LE', 'UTF-8')
      @@bytesize = :size
    end
    def setup
      @var = File.expand_path 'var', File.dirname(__FILE__)
      FileUtils.mkdir_p @var
      @data = File.expand_path 'data', File.dirname(__FILE__)
      FileUtils.mkdir_p @data
    end
    def teardown
      Spreadsheet.client_encoding = 'UTF-8'
      FileUtils.rm_r @var
    end
    def test_copy__identical__file_paths
      path = File.join @data, 'test_copy.xls'
      copy = File.join @data, 'test_copy1.xls'
      book = Spreadsheet.open path
      assert_instance_of Excel::Workbook, book
      book.write copy
      assert_equal File.read(path), File.read(copy)
    ensure
      File.delete copy if File.exist? copy
    end
    def test_empty_workbook
      path = File.join @data, 'test_empty.xls'
      book = Spreadsheet.open path
      assert_instance_of Excel::Workbook, book
      assert_equal 8, book.biff_version
      assert_equal 'Microsoft Excel 97/2000/XP', book.version_string
      enc = 'UTF-16LE'
      if defined? Encoding
        enc = Encoding.find enc
      end
      assert_equal enc, book.encoding
      assert_equal 21, book.formats.size
      assert_equal 4, book.fonts.size
      assert_equal 0, book.sst.size
      sheet = book.worksheet 0
      assert_equal 0, sheet.row_count
      assert_equal 0, sheet.column_count
      assert_nothing_raised do sheet.inspect end
    end
    def test_version_excel97__ooffice__utf16
      Spreadsheet.client_encoding = 'UTF-16LE'
      assert_equal 'UTF-16LE', Spreadsheet.client_encoding
      path = File.join @data, 'test_version_excel97.xls'
      book = Spreadsheet.open path
      assert_instance_of Excel::Workbook, book
      assert_equal 8, book.biff_version
      assert_equal @@iconv.iconv('Microsoft Excel 97/2000/XP'),
                   book.version_string
      enc = 'UTF-16LE'
      if defined? Encoding
        enc = Encoding.find enc
      end
      assert_equal enc, book.encoding
      assert_equal 25, book.formats.size
      assert_equal 5, book.fonts.size
      str1 = book.shared_string 0
      other = @@iconv.iconv('Shared String')
      assert_equal @@iconv.iconv('Shared String'), str1
      str2 = book.shared_string 1
      assert_equal @@iconv.iconv('Another Shared String'), str2
      str3 = book.shared_string 2
      long = @@iconv.iconv('1234567890 ' * 1000)
      if str3 != long
        long.size.times do |idx|
          len = idx.next
          if str3[0,len] != long[0,len]
            assert_equal long[idx - 5, 10], str3[idx - 5, 10], "in position #{idx}"
          end
        end
      end
      assert_equal long, str3
      str4 = book.shared_string 3
      long = @@iconv.iconv('9876543210 ' * 1000)
      if str4 != long
        long.size.times do |idx|
          len = idx.next
          if str4[0,len] != long[0,len]
            assert_equal long[idx - 5, 10], str4[idx - 5, 10], "in position #{idx}"
          end
        end
      end
      assert_equal long, str4
      sheet = book.worksheet 0
      assert_equal 11, sheet.row_count
      assert_equal 12, sheet.column_count
      useds = [0,0,0,0,0,0,0,1,0,0,11]
      unuseds = [2,2,1,1,1,2,1,11,1,2,12]
      sheet.each do |row|
        assert_equal useds.shift, row.first_used
        assert_equal unuseds.shift, row.first_unused
      end
      assert unuseds.empty?, "not all rows were visited in Spreadsheet#each"
      row = sheet.row 0
      assert_equal str1, row[0]
      assert_equal str1, sheet[0,0]
      assert_equal str1, sheet.cell(0,0)
      assert_equal str1, row[1]
      assert_equal str1, sheet[0,1]
      assert_equal str1, sheet.cell(0,1)
      row = sheet.row 1
      assert_equal str2, row[0]
      assert_equal str2, sheet[1,0]
      assert_equal str2, sheet.cell(1,0)
      assert_equal str2, row[1]
      assert_equal str2, sheet[1,1]
      assert_equal str2, sheet.cell(1,1)
      row = sheet.row 2
      assert_equal str3, row[0]
      assert_equal str3, sheet[2,0]
      assert_equal str3, sheet.cell(2,0)
      assert_nil row[1]
      assert_nil sheet[2,1]
      assert_nil sheet.cell(2,1)
      row = sheet.row 3
      assert_equal str4, row[0]
      assert_equal str4, sheet[3,0]
      assert_equal str4, sheet.cell(3,0)
      assert_nil row[1]
      assert_nil sheet[3,1]
      assert_nil sheet.cell(3,1)
      row = sheet.row 4
      assert_equal 0.25, row[0]
      assert_equal 0.25, sheet[4,0]
      assert_equal 0.25, sheet.cell(4,0)
      row = sheet.row 5
      assert_equal 0.75, row[0]
      assert_equal 0.75, sheet[5,0]
      assert_equal 0.75, sheet.cell(5,0)
      date = Date.new 1975, 8, 21
      assert_equal date, row[1]
      assert_equal date, sheet[5,1]
      assert_equal date, sheet.cell(5,1)
      row = sheet.row 6
      assert_equal 1, row[0]
      assert_equal 1, sheet[6,0]
      assert_equal 1, sheet.cell(6,0)
      row = sheet.row 7
      assert_nil row[0]
      assert_equal [1,2,3,4,5,6,7,8,9,0], row[1,10]
      assert_equal [1,2,3,4,5,6,7,8,9,0], sheet[7,1..10]
      assert_equal [1,2,3,4,5,6,7,8,9,0], sheet.cell(7,1..10)
      row = sheet.row 8
      assert_equal 0.0001, row[0]
      row = sheet.row 9
      assert_equal 0.00009, row[0]
      assert_equal :green, sheet.row(10).format(11).pattern_fg_color
    end
    def test_version_excel97__ooffice
      path = File.join @data, 'test_version_excel97.xls'
      book = Spreadsheet.open path
      assert_instance_of Excel::Workbook, book
      assert_equal 8, book.biff_version
      assert_equal 'Microsoft Excel 97/2000/XP', book.version_string
      enc = 'UTF-16LE'
      if defined? Encoding
        enc = Encoding.find enc
      end
      assert_equal enc, book.encoding
      assert_equal 25, book.formats.size
      assert_equal 5, book.fonts.size
      str1 = book.shared_string 0
      assert_equal 'Shared String', str1
      str2 = book.shared_string 1
      assert_equal 'Another Shared String', str2
      str3 = book.shared_string 2
      long = '1234567890 ' * 1000
      if str3 != long
        long.size.times do |idx|
          len = idx.next
          if str3[0,len] != long[0,len]
            assert_equal long[idx - 5, 10], str3[idx - 5, 10], "in position #{idx}"
          end
        end
      end
      assert_equal long, str3
      str4 = book.shared_string 3
      long = '9876543210 ' * 1000
      if str4 != long
        long.size.times do |idx|
          len = idx.next
          if str4[0,len] != long[0,len]
            assert_equal long[idx - 5, 10], str4[idx - 5, 10], "in position #{idx}"
          end
        end
      end
      assert_equal long, str4
      sheet = book.worksheet 0
      assert_equal 11, sheet.row_count
      assert_equal 12, sheet.column_count
      useds = [0,0,0,0,0,0,0,1,0,0,11]
      unuseds = [2,2,1,1,1,2,1,11,1,2,12]
      sheet.each do |row|
        assert_equal useds.shift, row.first_used
        assert_equal unuseds.shift, row.first_unused
      end
      assert unuseds.empty?, "not all rows were visited in Spreadsheet#each"
      row = sheet.row 0
      assert_equal str1, row[0]
      assert_equal str1, sheet[0,0]
      assert_equal str1, sheet.cell(0,0)
      assert_equal str1, row[1]
      assert_equal str1, sheet[0,1]
      assert_equal str1, sheet.cell(0,1)
      row = sheet.row 1
      assert_equal str2, row[0]
      assert_equal str2, sheet[1,0]
      assert_equal str2, sheet.cell(1,0)
      assert_equal str2, row[1]
      assert_equal str2, sheet[1,1]
      assert_equal str2, sheet.cell(1,1)
      row = sheet.row 2
      assert_equal str3, row[0]
      assert_equal str3, sheet[2,0]
      assert_equal str3, sheet.cell(2,0)
      assert_nil row[1]
      assert_nil sheet[2,1]
      assert_nil sheet.cell(2,1)
      row = sheet.row 3
      assert_equal str4, row[0]
      assert_equal str4, sheet[3,0]
      assert_equal str4, sheet.cell(3,0)
      assert_nil row[1]
      assert_nil sheet[3,1]
      assert_nil sheet.cell(3,1)
      row = sheet.row 4
      assert_equal 0.25, row[0]
      assert_equal 0.25, sheet[4,0]
      assert_equal 0.25, sheet.cell(4,0)
      row = sheet.row 5
      assert_equal 0.75, row[0]
      assert_equal 0.75, sheet[5,0]
      assert_equal 0.75, sheet.cell(5,0)
      date = Date.new 1975, 8, 21
      assert_equal date, row[1]
      assert_equal date, sheet[5,1]
      assert_equal date, sheet.cell(5,1)
      row = sheet.row 6
      assert_equal 1, row[0]
      assert_equal 1, sheet[6,0]
      assert_equal 1, sheet.cell(6,0)
      row = sheet.row 7
      assert_nil row[0]
      assert_equal [1,2,3,4,5,6,7,8,9,0], row[1,10]
      assert_equal [1,2,3,4,5,6,7,8,9,0], sheet[7,1..10]
      assert_equal [1,2,3,4,5,6,7,8,9,0], sheet.cell(7,1..10)
      row = sheet.row 8
      assert_equal 0.0001, row[0]
      row = sheet.row 9
      assert_equal 0.00009, row[0]
      link = row[1]
      assert_instance_of Link, link
      assert_equal 'Link-Text', link
      assert_equal 'http://scm.ywesee.com/spreadsheet', link.url
      assert_equal 'http://scm.ywesee.com/spreadsheet', link.href
    end
    def test_version_excel95__ooffice__utf16
      Spreadsheet.client_encoding = 'UTF-16LE'
      path = File.join @data, 'test_version_excel95.xls'
      book = Spreadsheet.open path
      assert_instance_of Excel::Workbook, book
      assert_equal 5, book.biff_version
      assert_equal @@iconv.iconv('Microsoft Excel 95'), book.version_string
      enc = 'WINDOWS-1252'
      if defined? Encoding
        enc = Encoding.find enc
      end
      assert_equal enc, book.encoding
      str1 = @@iconv.iconv('Shared String')
      str2 = @@iconv.iconv('Another Shared String')
      str3 = @@iconv.iconv(('1234567890 ' * 26)[0,255])
      str4 = @@iconv.iconv(('9876543210 ' * 26)[0,255])
      sheet = book.worksheet 0
      assert_equal 8, sheet.row_count
      assert_equal 11, sheet.column_count
      useds = [0,0,0,0,0,0,0,1]
      unuseds = [2,2,1,1,1,1,1,11]
      sheet.each do |row|
        assert_equal useds.shift, row.first_used
        assert_equal unuseds.shift, row.first_unused
      end
      assert unuseds.empty?, "not all rows were visited in Spreadsheet#each"
      row = sheet.row 0
      assert_equal str1, row[0]
      assert_equal str1, sheet[0,0]
      assert_equal str1, sheet.cell(0,0)
      assert_equal str1, row[1]
      assert_equal str1, sheet[0,1]
      assert_equal str1, sheet.cell(0,1)
      row = sheet.row 1
      assert_equal str2, row[0]
      assert_equal str2, sheet[1,0]
      assert_equal str2, sheet.cell(1,0)
      assert_equal str2, row[1]
      assert_equal str2, sheet[1,1]
      assert_equal str2, sheet.cell(1,1)
      row = sheet.row 2
      assert_equal 510, row[0].send(@@bytesize)
      assert_equal str3, row[0]
      assert_equal str3, sheet[2,0]
      assert_equal str3, sheet.cell(2,0)
      assert_nil row[1]
      assert_nil sheet[2,1]
      assert_nil sheet.cell(2,1)
      row = sheet.row 3
      assert_equal 510, row[0].send(@@bytesize)
      assert_equal str4, row[0]
      assert_equal str4, sheet[3,0]
      assert_equal str4, sheet.cell(3,0)
      assert_nil row[1]
      assert_nil sheet[3,1]
      assert_nil sheet.cell(3,1)
      row = sheet.row 4
      assert_equal 0.25, row[0]
      assert_equal 0.25, sheet[4,0]
      assert_equal 0.25, sheet.cell(4,0)
      row = sheet.row 5
      assert_equal 0.75, row[0]
      assert_equal 0.75, sheet[5,0]
      assert_equal 0.75, sheet.cell(5,0)
      row = sheet.row 6
      assert_equal 1, row[0]
      assert_equal 1, sheet[6,0]
      assert_equal 1, sheet.cell(6,0)
      row = sheet.row 7
      assert_nil row[0]
      assert_equal [1,2,3,4,5,6,7,8,9,0], row[1,10]
      assert_equal [1,2,3,4,5,6,7,8,9,0], sheet[7,1..10]
      assert_equal [1,2,3,4,5,6,7,8,9,0], sheet.cell(7,1..10)
    end
    def test_version_excel95__ooffice
      path = File.join @data, 'test_version_excel95.xls'
      book = Spreadsheet.open path
      assert_instance_of Excel::Workbook, book
      assert_equal 5, book.biff_version
      assert_equal 'Microsoft Excel 95', book.version_string
      enc = 'WINDOWS-1252'
      if defined? Encoding
        enc = Encoding.find enc
      end
      assert_equal enc, book.encoding
      str1 = 'Shared String'
      str2 = 'Another Shared String'
      str3 = ('1234567890 ' * 26)[0,255]
      str4 = ('9876543210 ' * 26)[0,255]
      sheet = book.worksheet 0
      assert_equal 8, sheet.row_count
      assert_equal 11, sheet.column_count
      useds = [0,0,0,0,0,0,0,1]
      unuseds = [2,2,1,1,1,1,1,11]
      sheet.each do |row|
        assert_equal useds.shift, row.first_used
        assert_equal unuseds.shift, row.first_unused
      end
      assert unuseds.empty?, "not all rows were visited in Spreadsheet#each"
      row = sheet.row 0
      assert_equal str1, row[0]
      assert_equal str1, sheet[0,0]
      assert_equal str1, sheet.cell(0,0)
      assert_equal str1, row[1]
      assert_equal str1, sheet[0,1]
      assert_equal str1, sheet.cell(0,1)
      row = sheet.row 1
      assert_equal str2, row[0]
      assert_equal str2, sheet[1,0]
      assert_equal str2, sheet.cell(1,0)
      assert_equal str2, row[1]
      assert_equal str2, sheet[1,1]
      assert_equal str2, sheet.cell(1,1)
      row = sheet.row 2
      assert_equal 255, row[0].send(@@bytesize)
      assert_equal str3, row[0]
      assert_equal str3, sheet[2,0]
      assert_equal str3, sheet.cell(2,0)
      assert_nil row[1]
      assert_nil sheet[2,1]
      assert_nil sheet.cell(2,1)
      row = sheet.row 3
      assert_equal 255, row[0].send(@@bytesize)
      assert_equal str4, row[0]
      assert_equal str4, sheet[3,0]
      assert_equal str4, sheet.cell(3,0)
      assert_nil row[1]
      assert_nil sheet[3,1]
      assert_nil sheet.cell(3,1)
      row = sheet.row 4
      assert_equal 0.25, row[0]
      assert_equal 0.25, sheet[4,0]
      assert_equal 0.25, sheet.cell(4,0)
      row = sheet.row 5
      assert_equal 0.75, row[0]
      assert_equal 0.75, sheet[5,0]
      assert_equal 0.75, sheet.cell(5,0)
      row = sheet.row 6
      assert_equal 1, row[0]
      assert_equal 1, sheet[6,0]
      assert_equal 1, sheet.cell(6,0)
      row = sheet.row 7
      assert_nil row[0]
      assert_equal [1,2,3,4,5,6,7,8,9,0], row[1,10]
      assert_equal [1,2,3,4,5,6,7,8,9,0], sheet[7,1..10]
      assert_equal [1,2,3,4,5,6,7,8,9,0], sheet.cell(7,1..10)
    end
    def test_version_excel5__ooffice
      path = File.join @data, 'test_version_excel5.xls'
      book = Spreadsheet.open path
      assert_instance_of Excel::Workbook, book
      assert_equal 5, book.biff_version
      assert_equal 'Microsoft Excel 95', book.version_string
      enc = 'WINDOWS-1252'
      if defined? Encoding
        enc = Encoding.find enc
      end
      assert_equal enc, book.encoding
      str1 = 'Shared String'
      str2 = 'Another Shared String'
      str3 = ('1234567890 ' * 26)[0,255]
      str4 = ('9876543210 ' * 26)[0,255]
      sheet = book.worksheet 0
      assert_equal 8, sheet.row_count
      assert_equal 11, sheet.column_count
      useds = [0,0,0,0,0,0,0,1]
      unuseds = [2,2,1,1,1,1,1,11]
      sheet.each do |row|
        assert_equal useds.shift, row.first_used
        assert_equal unuseds.shift, row.first_unused
      end
      assert unuseds.empty?, "not all rows were visited in Spreadsheet#each"
      row = sheet.row 0
      assert_equal str1, row[0]
      assert_equal str1, sheet[0,0]
      assert_equal str1, sheet.cell(0,0)
      assert_equal str1, row[1]
      assert_equal str1, sheet[0,1]
      assert_equal str1, sheet.cell(0,1)
      row = sheet.row 1
      assert_equal str2, row[0]
      assert_equal str2, sheet[1,0]
      assert_equal str2, sheet.cell(1,0)
      assert_equal str2, row[1]
      assert_equal str2, sheet[1,1]
      assert_equal str2, sheet.cell(1,1)
      row = sheet.row 2
      assert_equal 255, row[0].send(@@bytesize)
      assert_equal str3, row[0]
      assert_equal str3, sheet[2,0]
      assert_equal str3, sheet.cell(2,0)
      assert_nil row[1]
      assert_nil sheet[2,1]
      assert_nil sheet.cell(2,1)
      row = sheet.row 3
      assert_equal 255, row[0].send(@@bytesize)
      assert_equal str4, row[0]
      assert_equal str4, sheet[3,0]
      assert_equal str4, sheet.cell(3,0)
      assert_nil row[1]
      assert_nil sheet[3,1]
      assert_nil sheet.cell(3,1)
      row = sheet.row 4
      assert_equal 0.25, row[0]
      assert_equal 0.25, sheet[4,0]
      assert_equal 0.25, sheet.cell(4,0)
      row = sheet.row 5
      assert_equal 0.75, row[0]
      assert_equal 0.75, sheet[5,0]
      assert_equal 0.75, sheet.cell(5,0)
      row = sheet.row 6
      assert_equal 1, row[0]
      assert_equal 1, sheet[6,0]
      assert_equal 1, sheet.cell(6,0)
      row = sheet.row 7
      assert_nil row[0]
      assert_equal [1,2,3,4,5,6,7,8,9,0], row[1,10]
      assert_equal [1,2,3,4,5,6,7,8,9,0], sheet[7,1..10]
      assert_equal [1,2,3,4,5,6,7,8,9,0], sheet.cell(7,1..10)
    end
    def test_worksheets
      path = File.join @data, 'test_copy.xls'
      book = Spreadsheet.open path
      sheets = book.worksheets
      assert_equal 3, sheets.size
      sheet = book.worksheet 0
      assert_instance_of Excel::Worksheet, sheet
      assert_equal sheet, book.worksheet('Sheet1')
    end
    def test_worksheets__utf16
      Spreadsheet.client_encoding = 'UTF-16LE'
      path = File.join @data, 'test_copy.xls'
      book = Spreadsheet.open path
      sheets = book.worksheets
      assert_equal 3, sheets.size
      sheet = book.worksheet 0
      assert_instance_of Excel::Worksheet, sheet
      str = "S\000h\000e\000e\000t\0001\000"
      if RUBY_VERSION >= '1.9'
        str.force_encoding 'UTF-16LE' if name.respond_to?(:force_encoding)
      end
      assert_equal sheet, book.worksheet(str)
    end
    def test_read_datetime
      path = File.join @data, 'test_datetime.xls'
      book = Spreadsheet.open path
      assert_instance_of Excel::Workbook, book
      sheet = book.worksheet 0
      time = sheet[0,0]
      assert_equal 22, time.hour
      assert_equal 00, time.min
      assert_equal 00, time.sec
      time = sheet[1,0]
      assert_equal 1899, time.year
      assert_equal 12, time.month
      assert_equal 30, time.day
      assert_equal 22, time.hour
      assert_equal 30, time.min
      assert_equal 45, time.sec
      time = sheet[0,1]
      assert_equal 1899, time.year
      assert_equal 12, time.month
      assert_equal 31, time.day
      assert_equal 4, time.hour
      assert_equal 30, time.min
      assert_equal 45, time.sec
    end
    def test_change_encoding
      path = File.join @data, 'test_version_excel95.xls'
      book = Spreadsheet.open path
      assert_instance_of Excel::Workbook, book
      assert_equal 5, book.biff_version
      assert_equal 'Microsoft Excel 95', book.version_string
      enc = 'WINDOWS-1252'
      if defined? Encoding
        enc = Encoding.find enc
      end
      assert_equal enc, book.encoding
      enc = 'WINDOWS-1256'
      if defined? Encoding
        enc = Encoding.find enc
      end
      book.encoding = enc
      path = File.join @var, 'test_change_encoding.xls'
      book.write path
      assert_nothing_raised do book = Spreadsheet.open path end
      assert_equal enc, book.encoding
    end
    def test_change_cell
      path = File.join @data, 'test_version_excel97.xls'
      book = Spreadsheet.open path
      assert_instance_of Excel::Workbook, book
      assert_equal 8, book.biff_version
      assert_equal 'Microsoft Excel 97/2000/XP', book.version_string
      path = File.join @var, 'test_change_cell.xls'
      str1 = book.shared_string 0
      assert_equal 'Shared String', str1
      str2 = book.shared_string 1
      assert_equal 'Another Shared String', str2
      str3 = book.shared_string 2
      long = '1234567890 ' * 1000
      if str3 != long
        long.size.times do |idx|
          len = idx.next
          if str3[0,len] != long[0,len]
            assert_equal long[idx - 5, 10], str3[idx - 5, 10], "in position #{idx}"
          end
        end
      end
      assert_equal long, str3
      str4 = book.shared_string 3
      long = '9876543210 ' * 1000
      if str4 != long
        long.size.times do |idx|
          len = idx.next
          if str4[0,len] != long[0,len]
            assert_equal long[idx - 5, 10], str4[idx - 5, 10], "in position #{idx}"
          end
        end
      end
      assert_equal long, str4
      sheet = book.worksheet 0
      sheet[0,0] = 4
      row = sheet.row 1
      row[0] = 3
      book.write path
      assert_nothing_raised do book = Spreadsheet.open path end
      sheet = book.worksheet 0
      assert_equal 11, sheet.row_count
      assert_equal 12, sheet.column_count
      useds = [0,0,0,0,0,0,0,0,0,0,0]
      unuseds = [2,2,1,1,1,2,1,11,1,2,12]
      sheet.each do |rw|
        assert_equal useds.shift, rw.first_used
        assert_equal unuseds.shift, rw.first_unused
      end
      assert unuseds.empty?, "not all rows were visited in Spreadsheet#each"
      row = sheet.row 0
      assert_equal 4, row[0]
      assert_equal 4, sheet[0,0]
      assert_equal 4, sheet.cell(0,0)
      assert_equal str1, row[1]
      assert_equal str1, sheet[0,1]
      assert_equal str1, sheet.cell(0,1)
      row = sheet.row 1
      assert_equal 3, row[0]
      assert_equal 3, sheet[1,0]
      assert_equal 3, sheet.cell(1,0)
      assert_equal str2, row[1]
      assert_equal str2, sheet[1,1]
      assert_equal str2, sheet.cell(1,1)
      row = sheet.row 2
      assert_equal str3, row[0]
      assert_equal str3, sheet[2,0]
      assert_equal str3, sheet.cell(2,0)
      assert_nil row[1]
      assert_nil sheet[2,1]
      assert_nil sheet.cell(2,1)
      row = sheet.row 3
      assert_equal str4, row[0]
      assert_equal str4, sheet[3,0]
      assert_equal str4, sheet.cell(3,0)
      assert_nil row[1]
      assert_nil sheet[3,1]
      assert_nil sheet.cell(3,1)
      row = sheet.row 4
      assert_equal 0.25, row[0]
      assert_equal 0.25, sheet[4,0]
      assert_equal 0.25, sheet.cell(4,0)
      row = sheet.row 5
      assert_equal 0.75, row[0]
      assert_equal 0.75, sheet[5,0]
      assert_equal 0.75, sheet.cell(5,0)
      date = Date.new 1975, 8, 21
      assert_equal date, row[1]
      assert_equal date, sheet[5,1]
      assert_equal date, sheet.cell(5,1)
      row = sheet.row 6
      assert_equal 1, row[0]
      assert_equal 1, sheet[6,0]
      assert_equal 1, sheet.cell(6,0)
      row = sheet.row 7
      assert_nil row[0]
      assert_equal [1,2,3,4,5,6,7,8,9,0], row[1,10]
      assert_equal [1,2,3,4,5,6,7,8,9,0], sheet[7,1..10]
      assert_equal [1,2,3,4,5,6,7,8,9,0], sheet.cell(7,1..10)
      row = sheet.row 8
      assert_equal 0.0001, row[0]
      row = sheet.row 9
      assert_equal 0.00009, row[0]
    end
    def test_change_cell__complete_sst_rewrite
      path = File.join @data, 'test_version_excel97.xls'
      book = Spreadsheet.open path
      assert_instance_of Excel::Workbook, book
      assert_equal 8, book.biff_version
      assert_equal 'Microsoft Excel 97/2000/XP', book.version_string
      path = File.join @var, 'test_change_cell.xls'
      str1 = book.shared_string 0
      assert_equal 'Shared String', str1
      str2 = book.shared_string 1
      assert_equal 'Another Shared String', str2
      str3 = book.shared_string 2
      long = '1234567890 ' * 1000
      if str3 != long
        long.size.times do |idx|
          len = idx.next
          if str3[0,len] != long[0,len]
            assert_equal long[idx - 5, 10], str3[idx - 5, 10], "in position #{idx}"
          end
        end
      end
      assert_equal long, str3
      str4 = book.shared_string 3
      long = '9876543210 ' * 1000
      if str4 != long
        long.size.times do |idx|
          len = idx.next
          if str4[0,len] != long[0,len]
            assert_equal long[idx - 5, 10], str4[idx - 5, 10], "in position #{idx}"
          end
        end
      end
      assert_equal long, str4
      sheet = book.worksheet 0
      sheet[0,0] = 4
      str5 = 'A completely different String'
      sheet[0,1] = str5
      row = sheet.row 1
      row[0] = 3
      book.write path
      assert_nothing_raised do book = Spreadsheet.open path end
      assert_equal str5, book.shared_string(0)
      assert_equal str2, book.shared_string(1)
      assert_equal str3, book.shared_string(2)
      assert_equal str4, book.shared_string(3)
      sheet = book.worksheet 0
      assert_equal 11, sheet.row_count
      assert_equal 12, sheet.column_count
      useds = [0,0,0,0,0,0,0,0,0,0,0]
      unuseds = [2,2,1,1,1,2,1,11,1,2,12]
      sheet.each do |rw|
        assert_equal useds.shift, rw.first_used
        assert_equal unuseds.shift, rw.first_unused
      end
      assert unuseds.empty?, "not all rows were visited in Spreadsheet#each"
      row = sheet.row 0
      assert_equal 4, row[0]
      assert_equal 4, sheet[0,0]
      assert_equal 4, sheet.cell(0,0)
      assert_equal str5, row[1]
      assert_equal str5, sheet[0,1]
      assert_equal str5, sheet.cell(0,1)
      row = sheet.row 1
      assert_equal 3, row[0]
      assert_equal 3, sheet[1,0]
      assert_equal 3, sheet.cell(1,0)
      assert_equal str2, row[1]
      assert_equal str2, sheet[1,1]
      assert_equal str2, sheet.cell(1,1)
      row = sheet.row 2
      assert_equal str3, row[0]
      assert_equal str3, sheet[2,0]
      assert_equal str3, sheet.cell(2,0)
      assert_nil row[1]
      assert_nil sheet[2,1]
      assert_nil sheet.cell(2,1)
      row = sheet.row 3
      assert_equal str4, row[0]
      assert_equal str4, sheet[3,0]
      assert_equal str4, sheet.cell(3,0)
      assert_nil row[1]
      assert_nil sheet[3,1]
      assert_nil sheet.cell(3,1)
      row = sheet.row 4
      assert_equal 0.25, row[0]
      assert_equal 0.25, sheet[4,0]
      assert_equal 0.25, sheet.cell(4,0)
      row = sheet.row 5
      assert_equal 0.75, row[0]
      assert_equal 0.75, sheet[5,0]
      assert_equal 0.75, sheet.cell(5,0)
      date = Date.new 1975, 8, 21
      assert_equal date, row[1]
      assert_equal date, sheet[5,1]
      assert_equal date, sheet.cell(5,1)
      row = sheet.row 6
      assert_equal 1, row[0]
      assert_equal 1, sheet[6,0]
      assert_equal 1, sheet.cell(6,0)
      row = sheet.row 7
      assert_nil row[0]
      assert_equal [1,2,3,4,5,6,7,8,9,0], row[1,10]
      assert_equal [1,2,3,4,5,6,7,8,9,0], sheet[7,1..10]
      assert_equal [1,2,3,4,5,6,7,8,9,0], sheet.cell(7,1..10)
      row = sheet.row 8
      assert_equal 0.0001, row[0]
      row = sheet.row 9
      assert_equal 0.00009, row[0]
    end
    def test_write_to_stringio
      book = Spreadsheet::Excel::Workbook.new
      sheet = book.create_worksheet :name => 'My Worksheet'
      sheet[0,0] = 'my cell'
      data = StringIO.new ''
      assert_nothing_raised do
        book.write data
      end
      assert_nothing_raised do
        book = Spreadsheet.open data
      end
      assert_instance_of Spreadsheet::Excel::Workbook, book
      assert_equal 1, book.worksheets.size
      sheet = book.worksheet 0
      assert_equal 'My Worksheet', sheet.name
      assert_equal 'my cell', sheet[0,0]
    end
    def test_write_new_workbook
      book = Spreadsheet::Workbook.new
      path = File.join @var, 'test_write_workbook.xls'
      sheet1 = book.create_worksheet
      str1 = 'My Shared String'
      str2 = 'Another Shared String'
      assert_equal 1, (str1.size + str2.size) % 2, 
        "str3 should start at an odd offset to test splitting of wide strings"
      str3 = '–––––––––– ' * 1000
      str4 = '1234567890 ' * 1000
      fmt1 = Format.new :italic => true, :color => :blue
      sheet1.format_column 1, fmt1, :width => 20
      fmt2 = Format.new(:weight => :bold, :color => :yellow)
      sheet1.format_column 2, fmt2
      sheet1.format_column 3, Format.new(:weight => :bold, :color => :red)
      sheet1.format_column 6..9, fmt1
      sheet1.format_column [4,5,7], fmt2
      sheet1.row(0).height = 20
      sheet1[0,0] = str1
      sheet1.row(0).push str1
      sheet1.row(1).concat [str2, str2]
      sheet1[2,0] = str3
      sheet1[3,0] = str4
      fmt = Format.new :color => 'red'
      sheet1[4,0] = 0.25
      sheet1.row(4).set_format 0, fmt
      fmt = Format.new :color => 'aqua'
      sheet1[5,0] = 0.75
      sheet1.row(5).set_format 0, fmt
      link = Link.new 'http://scm.ywesee.com/?p=spreadsheet;a=summary',
                      'The Spreadsheet GitWeb'
      sheet1[5,1] = link
      sheet1[6,0] = 1
      fmt = Format.new :color => 'green'
      sheet1.row(6).set_format 0, fmt
      sheet1[6,1] = Date.new 2008, 10, 10
      sheet1[6,2] = Date.new 2008, 10, 12
      fmt = Format.new :number_format => 'D.M.YY'
      sheet1.row(6).set_format 1, fmt
      sheet1.update_row 7, nil, 1, 2, 3, 4, 5, 6, 7, 8, 9, 0
      sheet1[8,0] = 0.0005
      sheet1[8,1] = 0.005
      sheet1[8,2] = 0.05
      sheet1[8,3] = 10.5
      sheet1[8,4] = 1.05
      sheet1[8,5] = 100.5
      sheet1[8,6] = 10.05
      sheet1[8,7] = 1.005
      sheet1[9,0] = 100.5
      sheet1[9,1] = 10.05
      sheet1[9,2] = 1.005
      sheet1[9,3] = 1000.5
      sheet1[9,4] = 100.05
      sheet1[9,5] = 10.005
      sheet1[9,6] = 1.0005
      sheet1[10,0] = 10000.5
      sheet1[10,1] = 1000.05
      sheet1[10,2] = 100.005
      sheet1[10,3] = 10.0005
      sheet1[10,4] = 1.00005
      sheet1.insert_row 9, ['a', 'b', 'c']
      assert_equal 'a', sheet1[9,0]
      assert_equal 'b', sheet1[9,1]
      assert_equal 'c', sheet1[9,2]
      sheet1.delete_row 9
      row = sheet1.row(11)
      row.height = 40
      row.push 'x'
      row.pop
      sheet2 = book.create_worksheet :name => 'my name'
      book.write path
      Spreadsheet.client_encoding = 'UTF-16LE'
      str1 = @@iconv.iconv str1
      str2 = @@iconv.iconv str2
      str3 = @@iconv.iconv str3
      str4 = @@iconv.iconv str4
      assert_nothing_raised do book = Spreadsheet.open path end
      if RUBY_VERSION >= '1.9'
        assert_equal 'UTF-16LE', book.encoding.name
      else
        assert_equal 'UTF-16LE', book.encoding
      end
      assert_equal str1, book.shared_string(0)
      assert_equal str2, book.shared_string(1)
      test = nil
      assert_nothing_raised "I've probably split a two-byte-character" do
        test = book.shared_string 2
      end
      if test != str3
        str3.size.times do |idx|
          len = idx.next
          if test[0,len] != str3[0,len]
            assert_equal str3[idx - 5, 10], test[idx - 5, 10], "in position #{idx}"
          end
        end
      end
      assert_equal str3, test
      test = book.shared_string 3
      if test != str4
        str4.size.times do |idx|
          len = idx.next
          if test[0,len] != str4[0,len]
            assert_equal str4[idx - 5, 10], test[idx - 5, 10], "in position #{idx}"
          end
        end
      end
      assert_equal str4, test
      assert_equal 2, book.worksheets.size
      sheet = book.worksheets.first
      assert_instance_of Spreadsheet::Excel::Worksheet, sheet
      name = "W\000o\000r\000k\000s\000h\000e\000e\000t\0001\000"
      name.force_encoding 'UTF-16LE' if name.respond_to?(:force_encoding)
      assert_equal name, sheet.name
      assert_not_nil sheet.offset
      assert_not_nil col = sheet.column(1)
      assert_equal true, col.default_format.font.italic?
      assert_equal :blue, col.default_format.font.color
      assert_equal 20, col.width
      row = sheet.row 0
      assert_equal col.default_format, row.format(1)
      assert_equal 20, row.height
      assert_equal str1, row[0]
      assert_equal str1, sheet[0,0]
      assert_equal str1, sheet.cell(0,0)
      assert_equal str1, row[1]
      assert_equal str1, sheet[0,1]
      assert_equal str1, sheet.cell(0,1)
      row = sheet.row 1
      assert_equal str2, row[0]
      assert_equal str2, sheet[1,0]
      assert_equal str2, sheet.cell(1,0)
      assert_equal str2, row[1]
      assert_equal str2, sheet[1,1]
      assert_equal str2, sheet.cell(1,1)
      row = sheet.row 2
      assert_equal str3, row[0]
      assert_equal str3, sheet[2,0]
      assert_equal str3, sheet.cell(2,0)
      assert_nil row[1]
      assert_nil sheet[2,1]
      assert_nil sheet.cell(2,1)
      row = sheet.row 3
      assert_equal str4, row[0]
      assert_equal str4, sheet[3,0]
      assert_equal str4, sheet.cell(3,0)
      assert_nil row[1]
      assert_nil sheet[3,1]
      assert_nil sheet.cell(3,1)
      row = sheet.row 4
      assert_equal :red, row.format(0).font.color
      assert_equal 0.25, row[0]
      assert_equal 0.25, sheet[4,0]
      assert_equal 0.25, sheet.cell(4,0)
      row = sheet.row 5
      assert_equal :cyan, row.format(0).font.color
      assert_equal 0.75, row[0]
      assert_equal 0.75, sheet[5,0]
      assert_equal 0.75, sheet.cell(5,0)
      link = row[1]
      assert_instance_of Link, link
      url = @@iconv.iconv 'http://scm.ywesee.com/?p=spreadsheet;a=summary'
      assert_equal @@iconv.iconv('The Spreadsheet GitWeb'), link
      assert_equal url, link.url
      row = sheet.row 6
      assert_equal :green, row.format(0).font.color
      assert_equal 1, row[0]
      assert_equal 1, sheet[6,0]
      assert_equal 1, sheet.cell(6,0)
      assert_equal @@iconv.iconv('D.M.YY'), row.format(1).number_format
      date = Date.new 2008, 10, 10
      assert_equal date, row[1]
      assert_equal date, sheet[6,1]
      assert_equal date, sheet.cell(6,1)
      assert_equal @@iconv.iconv('DD.MM.YYYY'), row.format(2).number_format
      date = Date.new 2008, 10, 12
      assert_equal date, row[2]
      assert_equal date, sheet[6,2]
      assert_equal date, sheet.cell(6,2)
      row = sheet.row 7
      assert_nil row[0]
      assert_equal [1,2,3,4,5,6,7,8,9,0], row[1,10]
      assert_equal [1,2,3,4,5,6,7,8,9,0], sheet[7,1..10]
      assert_equal [1,2,3,4,5,6,7,8,9,0], sheet.cell(7,1..10)
      assert_equal 0.0005, sheet1[8,0]
      assert_equal 0.005, sheet1[8,1]
      assert_equal 0.05, sheet1[8,2]
      assert_equal 10.5, sheet1[8,3]
      assert_equal 1.05, sheet1[8,4]
      assert_equal 100.5, sheet1[8,5]
      assert_equal 10.05, sheet1[8,6]
      assert_equal 1.005, sheet1[8,7]
      assert_equal 100.5, sheet1[9,0]
      assert_equal 10.05, sheet1[9,1]
      assert_equal 1.005, sheet1[9,2]
      assert_equal 1000.5, sheet1[9,3]
      assert_equal 100.05, sheet1[9,4]
      assert_equal 10.005, sheet1[9,5]
      assert_equal 1.0005, sheet1[9,6]
      assert_equal 10000.5, sheet1[10,0]
      assert_equal 1000.05, sheet1[10,1]
      assert_equal 100.005, sheet1[10,2]
      assert_equal 10.0005, sheet1[10,3]
      assert_equal 1.00005, sheet1[10,4]
      assert_equal 40, sheet1.row(11).height
      assert_instance_of Spreadsheet::Excel::Worksheet, sheet
      sheet = book.worksheets.last
      name = "m\000y\000 \000n\000a\000m\000e\000"
      name.force_encoding 'UTF-16LE' if name.respond_to?(:force_encoding)
      assert_equal name, sheet.name
      assert_not_nil sheet.offset
    end
    def test_write_new_workbook__utf16
      Spreadsheet.client_encoding = 'UTF-16LE'
      book = Spreadsheet::Workbook.new
      path = File.join @var, 'test_write_workbook.xls'
      sheet1 = book.create_worksheet
      str1 = @@iconv.iconv 'Shared String'
      str2 = @@iconv.iconv 'Another Shared String'
      str3 = @@iconv.iconv('1234567890 ' * 1000)
      str4 = @@iconv.iconv('9876543210 ' * 1000)
      fmt = Format.new :italic => true, :color => :blue
      sheet1.format_column 1, fmt, :width => 20
      sheet1[0,0] = str1
      sheet1.row(0).push str1
      sheet1.row(1).concat [str2, str2]
      sheet1[2,0] = str3
      sheet1[3,0] = str4
      fmt = Format.new :color => 'red'
      sheet1[4,0] = 0.25
      sheet1.row(4).set_format 0, fmt
      fmt = Format.new :color => 'aqua'
      sheet1[5,0] = 0.75
      sheet1.row(5).set_format 0, fmt
      sheet1[6,0] = 1
      fmt = Format.new :color => 'green'
      sheet1.row(6).set_format 0, fmt
      sheet1[6,1] = Date.new 2008, 10, 10
      sheet1[6,2] = Date.new 2008, 10, 12
      fmt = Format.new :number_format => @@iconv.iconv("DD.MM.YYYY")
      sheet1.row(6).set_format 1, fmt
      sheet1.update_row 7, nil, 1, 2, 3, 4, 5, 6, 7, 8, 9, 0
      sheet1.row(8).default_format = fmt
      sheet1[8,0] = @@iconv.iconv 'formatted when empty'
      sheet2 = book.create_worksheet :name => @@iconv.iconv("my name")
      book.write path
      Spreadsheet.client_encoding = 'UTF-8'
      str1 = 'Shared String'
      str2 = 'Another Shared String'
      str3 = '1234567890 ' * 1000
      str4 = '9876543210 ' * 1000
      assert_nothing_raised do book = Spreadsheet.open path end
      if RUBY_VERSION >= '1.9'
        assert_equal 'UTF-16LE', book.encoding.name
      else
        assert_equal 'UTF-16LE', book.encoding
      end
      assert_equal str1, book.shared_string(0)
      assert_equal str2, book.shared_string(1)
      test = book.shared_string 2
      if test != str3
        str3.size.times do |idx|
          len = idx.next
          if test[0,len] != str3[0,len]
            assert_equal str3[idx - 5, 10], test[idx - 5, 10], "in position #{idx}"
          end
        end
      end
      assert_equal str3, test
      test = book.shared_string 3
      if test != str4
        str4.size.times do |idx|
          len = idx.next
          if test[0,len] != str4[0,len]
            assert_equal str4[idx - 5, 10], test[idx - 5, 10], "in position #{idx}"
          end
        end
      end
      assert_equal str4, test
      assert_equal 2, book.worksheets.size
      sheet = book.worksheets.first
      assert_instance_of Spreadsheet::Excel::Worksheet, sheet
      assert_equal "Worksheet1", sheet.name
      assert_not_nil sheet.offset
      assert_not_nil col = sheet.column(1)
      assert_equal true, col.default_format.font.italic?
      assert_equal :blue, col.default_format.font.color
      row = sheet.row 0
      assert_equal col.default_format, row.format(1)
      assert_equal str1, row[0]
      assert_equal str1, sheet[0,0]
      assert_equal str1, sheet.cell(0,0)
      assert_equal str1, row[1]
      assert_equal str1, sheet[0,1]
      assert_equal str1, sheet.cell(0,1)
      row = sheet.row 1
      assert_equal str2, row[0]
      assert_equal str2, sheet[1,0]
      assert_equal str2, sheet.cell(1,0)
      assert_equal str2, row[1]
      assert_equal str2, sheet[1,1]
      assert_equal str2, sheet.cell(1,1)
      row = sheet.row 2
      assert_equal str3, row[0]
      assert_equal str3, sheet[2,0]
      assert_equal str3, sheet.cell(2,0)
      assert_nil row[1]
      assert_nil sheet[2,1]
      assert_nil sheet.cell(2,1)
      row = sheet.row 3
      assert_equal str4, row[0]
      assert_equal str4, sheet[3,0]
      assert_equal str4, sheet.cell(3,0)
      assert_nil row[1]
      assert_nil sheet[3,1]
      assert_nil sheet.cell(3,1)
      row = sheet.row 4
      assert_equal :red, row.format(0).font.color
      assert_equal 0.25, row[0]
      assert_equal 0.25, sheet[4,0]
      assert_equal 0.25, sheet.cell(4,0)
      row = sheet.row 5
      assert_equal :cyan, row.format(0).font.color
      assert_equal 0.75, row[0]
      assert_equal 0.75, sheet[5,0]
      assert_equal 0.75, sheet.cell(5,0)
      row = sheet.row 6
      assert_equal :green, row.format(0).font.color
      assert_equal 1, row[0]
      assert_equal 1, sheet[6,0]
      assert_equal 1, sheet.cell(6,0)
      assert_equal 'DD.MM.YYYY', row.format(1).number_format
      date = Date.new 2008, 10, 10
      assert_equal date, row[1]
      assert_equal date, sheet[6,1]
      assert_equal date, sheet.cell(6,1)
      assert_equal 'DD.MM.YYYY', row.format(2).number_format
      date = Date.new 2008, 10, 12
      assert_equal date, row[2]
      assert_equal date, sheet[6,2]
      assert_equal date, sheet.cell(6,2)
      row = sheet.row 7
      assert_nil row[0]
      assert_equal [1,2,3,4,5,6,7,8,9,0], row[1,10]
      assert_equal [1,2,3,4,5,6,7,8,9,0], sheet[7,1..10]
      assert_equal [1,2,3,4,5,6,7,8,9,0], sheet.cell(7,1..10)
      row = sheet.row 8
      assert_equal 'formatted when empty', row[0]
      assert_not_nil row.default_format
      assert_instance_of Spreadsheet::Excel::Worksheet, sheet
      sheet = book.worksheets.last
      assert_equal "my name",
                   sheet.name
      assert_not_nil sheet.offset
    end
    def test_template
      template = File.join @data, 'test_copy.xls'
      output = File.join @var, 'test_template.xls'
      book = Spreadsheet.open template
      sheet1 = book.worksheet 0
      sheet1.row(4).replace [ 'Daniel J. Berger', 'U.S.A.',
        'Author of original code for Spreadsheet::Excel' ]
      book.write output
      assert_nothing_raised do
        book = Spreadsheet.open output
      end
      sheet = book.worksheet 0
      row = sheet.row(4)
      assert_equal 'Daniel J. Berger', row[0]
    end
    def test_bignum
      smallnum = 0x1fffffff
      bignum = smallnum + 1
      book = Spreadsheet::Workbook.new
      sheet = book.create_worksheet
      sheet[0,0] = bignum
      sheet[1,0] = -bignum
      sheet[0,1] = smallnum
      sheet[1,1] = -smallnum
      sheet[0,2] = bignum - 0.1
      sheet[1,2] = -bignum - 0.1
      sheet[0,3] = smallnum - 0.1
      sheet[1,3] = -smallnum - 0.1
      path = File.join @var, 'test_big-number.xls'
      book.write path
      assert_nothing_raised do
        book = Spreadsheet.open path
      end
      assert_equal bignum, book.worksheet(0)[0,0]
      assert_equal(-bignum, book.worksheet(0)[1,0])
      assert_equal smallnum, book.worksheet(0)[0,1]
      assert_equal(-smallnum, book.worksheet(0)[1,1])
      assert_equal bignum - 0.1, book.worksheet(0)[0,2]
      assert_equal(-bignum - 0.1, book.worksheet(0)[1,2])
      assert_equal smallnum - 0.1, book.worksheet(0)[0,3]
      assert_equal(-smallnum - 0.1, book.worksheet(0)[1,3])
    end
    def test_bigfloat
      # reported in http://rubyforge.org/tracker/index.php?func=detail&aid=24119&group_id=678&atid=2677
      bigfloat = 10000000.0
      book = Spreadsheet::Workbook.new
      sheet = book.create_worksheet
      sheet[0,0] = bigfloat
      sheet[0,1] = bigfloat + 0.1
      sheet[0,2] = bigfloat - 0.1
      sheet[1,0] = -bigfloat
      sheet[1,1] = -bigfloat + 0.1
      sheet[1,2] = -bigfloat - 0.1
      path = File.join @var, 'test_big-float.xls'
      book.write path
      assert_nothing_raised do
        book = Spreadsheet.open path
      end
      sheet = book.worksheet(0)
      assert_equal bigfloat, sheet[0,0]
      assert_equal bigfloat + 0.1, sheet[0,1]
      assert_equal bigfloat - 0.1, sheet[0,2]
      assert_equal(-bigfloat, sheet[1,0])
      assert_equal(-bigfloat + 0.1, sheet[1,1])
      assert_equal(-bigfloat - 0.1, sheet[1,2])
    end
    def test_datetime__off_by_one
      # reported in http://rubyforge.org/tracker/index.php?func=detail&aid=24414&group_id=678&atid=2677
      datetime1 = DateTime.new(2008)
      datetime2 = DateTime.new(2008, 1, 1, 1, 0, 1)
      date1 = Date.new(2008)
      date2 = Date.new(2009)
      book = Spreadsheet::Workbook.new
      sheet = book.create_worksheet
      sheet[0,0] = datetime1
      sheet[0,1] = datetime2
      sheet[1,0] = date1
      sheet[1,1] = date2
      path = File.join @var, 'test_datetime.xls'
      book.write path
      assert_nothing_raised do
        book = Spreadsheet.open path
      end
      sheet = book.worksheet(0)
      assert_equal datetime1, sheet[0,0]
      assert_equal datetime2, sheet[0,1]
      assert_equal date1, sheet[1,0]
      assert_equal date2, sheet[1,1]
      assert_equal date1, sheet.row(0).date(0)
      assert_equal datetime1, sheet.row(1).datetime(0)
    end
    def test_sharedfmla
      path = File.join @data, 'test_formula.xls'
      book = Spreadsheet.open path
      assert_instance_of Excel::Workbook, book
      sheet = book.worksheet 0
      64.times do |idx|
        assert_equal '5026', sheet[idx.next, 2].value
      end
    end
    def test_missing_row_op
      path = File.join @data, 'test_missing_row.xls'
      book = Spreadsheet.open path
      assert_instance_of Excel::Workbook, book
      sheet = book.worksheet 0
      assert_not_nil sheet[1,0]
      assert_not_nil sheet[2,1]
    end
    def test_changes
      path = File.join @data, 'test_changes.xls'
      book = Spreadsheet.open path
      assert_instance_of Excel::Workbook, book
      sheet = book.worksheet 1
      sheet[20,0] = 'Ciao Mundo!'
      target = File.join @var, 'test_changes.xls'
      assert_nothing_raised do book.write target end
    end
  end
end
