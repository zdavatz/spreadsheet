require 'spreadsheet/encodings'

module Spreadsheet
  ##
  # The Note class is a Subclass of String and represents a comment/note/annotation
  # someone made to a cell.
  #
  #
  # Interesting Attributes
  # #author  :: The name of the author who wrote the note
  class Note < String
    include Encodings
    attr_accessor :author, :length, :objID, :row, :col
    def initialize
      super ''
      @author = nil
      @length = 0
      @objID  = nil
      @row    = -1
      @col    = -1
    end
  end
end
