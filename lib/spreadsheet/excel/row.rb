require 'date'
require 'spreadsheet/row'

module Spreadsheet
  module Excel
##
# Excel-specific Row methods
class Row < Spreadsheet::Row
  ##
  # The Excel date calculation erroneously assumes that 1900 is a leap-year. All
  # Dates after 28.2.1900 are off by one.
  LEAP_ERROR = Date.new 1900, 2, 28
  ##
  # Force convert the cell at _idx_ to a Date
  def date idx
    _date at(idx)
  end
  ##
  # Force convert the cell at _idx_ to a DateTime
  def datetime idx
    _datetime at(idx)
  end
  ##
  # Access data in this Row like you would in an Array. If a cell is formatted
  # as a Date or DateTime, the decoded Date or DateTime value is returned.
  def [] idx, len=nil
    if len
      idx = idx...(idx+len)
    end
    if idx.is_a? Range
      data = []
      idx.each do |i|
        data.push enriched_data(i, at(i))
      end
      data
    else
      enriched_data idx, at(idx)
    end
  end
  private
  def _date data # :nodoc:
    return data if data.is_a?(Date)
    date = @worksheet.date_base + data.to_i
    if date > LEAP_ERROR
      date -= 1
    end
    date
  end
  def _datetime data # :nodoc:
    return data if data.is_a?(DateTime)
    date = _date data
    DateTime.new(date.year, date.month, date.day) + (data.to_f % 1)
  end
  def enriched_data idx, data # :nodoc:
    res = nil
    if link = @worksheet.links[[@idx, idx]]
      res = link
    elsif data.is_a?(Numeric) && fmt = format(idx)
      res = if fmt.datetime? || fmt.time?
              _datetime data
            elsif fmt.date?
              _date data
            end
    end
    res || data
  end
end
  end
end
