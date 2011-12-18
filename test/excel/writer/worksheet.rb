#!/usr/bin/env ruby
# Excel::Writer::TestWorksheet -- Spreadheet -- 21.11.2007 -- hwyss@ywesee.com

require 'test/unit'
require 'spreadsheet/excel/writer/worksheet'

module Spreadsheet
  module Excel
    module Writer
class TestWorksheet < Test::Unit::TestCase
  def test_need_number
    sheet = Worksheet.new nil, nil
    assert_equal false, sheet.need_number?(10)
    assert_equal false, sheet.need_number?(114.55)
    assert_equal false, sheet.need_number?(0.1)
    assert_equal false, sheet.need_number?(0.01)
    assert_equal false, sheet.need_number?(0 / 0.0) # NaN
    assert_equal true, sheet.need_number?(0.001)
    assert_equal true, sheet.need_number?(10000000.0)
  end
  def test_write_merged_cells
    # top/bottom/left/right cell of test range 1/2
    r1t, r1b, r1l, r1r = 0, 0, 0, 1
    r2t, r2b, r2l, r2r = 1, 2, 0, 0
    book = Spreadsheet::Workbook.new
    sheet = book.create_worksheet
    sheet.merge_cells(r1t, r1l, r1b, r1r)
    sheet.merge_cells(r2t, r2l, r2b, r2r)
    assert_equal [[r1t, r1b, r1l, r1r], [r2t, r2b, r2l, r2r]], sheet.merged_cells
    io = StringIO.new
    book.write(io)
    book2 = Spreadsheet.open(io)
    sheet2 = book2.worksheet(0)
    sheet2[0,0] # trigger read_worksheet
    assert_equal [[r1t, r1b, r1l, r1r], [r2t, r2b, r2l, r2r]], sheet2.merged_cells
  end
end
    end
  end
end
