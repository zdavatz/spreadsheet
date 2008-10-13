module Spreadsheet
  module Excel
    class Reader
##
# This Module collects reader methods such as read_string that are specific to
# Biff5.  This Module is likely to be expanded as Support for older Versions
# of Excel grows.
module Biff5
  ##
  # Read a String of 8-bit Characters
  def read_string work, count_length=1
    # Offset    Size  Contents
    #      0  1 or 2  Length of the string (character count, ln)
    # 1 or 2      ln  Character array (8-bit characters)
    fmt = count_length == 1 ? 'C' : 'v'
    length, = work.unpack fmt
    work[count_length, length]
  end
end
    end
  end
end
