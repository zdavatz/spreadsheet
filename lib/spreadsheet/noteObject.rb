require "spreadsheet/encodings"

module Spreadsheet
  ##
  # The NoteObject class is made to handle the text output from the
  # object, txo, continue records which contain a comment's text record.
  #
  #
  class NoteObject
    include Encodings
    attr_accessor :obj_id, :text
    def initialize
      @obj_id = -1
      @text = ""
    end
  end
end
