require "spreadsheet/encodings"

module Spreadsheet
  ##
  # The Note class is a Subclass of String and represents a comment/note/annotation
  # someone made to a cell.
  #
  #
  # Interesting Attributes
  # #author  :: The name of the author who wrote the note
  class Note
    include Encodings
    attr_accessor :author, :length, :obj_id, :row, :col, :text
    def initialize
      @author = nil
      @length = 0
      @obj_id = nil
      @row = -1
      @col = -1
      @text = ""
    end
  end
end
