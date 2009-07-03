# encoding: utf-8
require 'spreadsheet/datatypes'
require 'spreadsheet/encodings'
require 'spreadsheet/font'

module Spreadsheet
  ##
  # Formatting data
  class Format
    include Spreadsheet::Datatypes
    include Spreadsheet::Encodings
    ##
    # You can set the following boolean attributes:
    # #cross_down::       Draws a Line from the top-left to the bottom-right
    #                     corner of a cell.
    # #cross_up::         Draws a Line from the bottom-left to the top-right
    #                     corner of a cell.
    # #hidden::           The cell is hidden.
    # #locked::           The cell is locked.
    # #merge_range::      The cell is in a merged range.
    # #shrink::           Shrink the contents to fit the cell.
    # #text_justlast::    Force the last line of a cell to be justified. This
    #                     probably makes sense if horizontal_align = :justify
    # #left::             Draw a border to the left of the cell.
    # #right::            Draw a border to the right of the cell.
    # #top::              Draw a border at the top of the cell.
    # #bottom::           Draw a border at the bottom of the cell.
    # #rotation_stacked:: Characters in the cell are stacked on top of each
    #                     other. Excel will ignore other rotation values if
    #                     this is set.
    boolean :cross_down, :cross_up, :hidden, :locked,
            :merge_range, :shrink, :text_justlast, :text_wrap, :left, :right,
            :top, :bottom, :rotation_stacked
    ##
    # Color attributes
    colors  :bottom_color, :top_color, :left_color, :right_color,
            :pattern_fg_color, :pattern_bg_color,
            :diagonal_color
    ##
    # Text direction
    # Valid values: :context, :left_to_right, :right_to_left
    # Default:      :context
    enum :text_direction, :context, :left_to_right, :right_to_left,
         :left_to_right => [:ltr, :l2r],
         :right_to_left => [:rtl, :r2l]
    alias :reading_order  :text_direction
    alias :reading_order= :text_direction=
    ##
    # Indentation level
    enum :indent_level, 0, Integer
    alias :indent  :indent_level
    alias :indent= :indent_level=
    ##
    # Horizontal alignment
    # Valid values: :default, :left, :center, :right, :fill, :justify, :merge,
    #               :distributed
    # Default:      :default
    enum :horizontal_align, :default, :left, :center, :right, :fill, :justify,
                            :merge, :distributed,
         :center      => :centre,
         :merge       => [ :center_across, :centre_across ],
         :distributed => :equal_space
    ##
    # Vertical alignment
    # Valid values: :bottom, :top, :middle, :justify, :distributed
    # Default:      :bottom
    enum :vertical_align, :bottom, :top, :middle, :justify, :distributed,
         :distributed => [:vdistributed, :vequal_space, :equal_space],
         :justify     => :vjustify,
         :middle      => [:vcenter, :vcentre, :center, :centre]
    attr_accessor :font, :number_format, :name, :pattern, :used_merge
    ##
    # Text rotation
    attr_reader :rotation
    def initialize opts={}
      @font             = Font.new client("Arial", 'UTF-8'), :family => :swiss
      @number_format    = client 'GENERAL', 'UTF-8'
      @rotation         = 0
      @pattern          = 0
      @bottom_color     = :builtin_black
      @top_color        = :builtin_black
      @left_color       = :builtin_black
      @right_color      = :builtin_black
      @diagonal_color   = :builtin_black
      @pattern_fg_color = :border
      @pattern_bg_color = :pattern_bg
      # Temp code to prevent merged formats in non-merged cells.
      @used_merge    = 0
      opts.each do |key, val|
        writer = "#{key}="
        if @font.respond_to? writer
          @font.send writer, val
        else
          self.send writer, val
        end
      end
      yield self if block_given?
    end
    ##
    # Combined method for both horizontal and vertical alignment. Sets the
    # first valid value (e.g. Format#align = :justify only sets the horizontal
    # alignment. Use one of the aliases prefixed with :v if you need to
    # disambiguate.)
    #
    # This is essentially a backward-compatibility method and may be removed at
    # some point in the future.
    def align= location
      self.horizontal_align = location
    rescue ArgumentError
      self.vertical_align = location rescue ArgumentError
    end
    ##
    # Returns an Array containing the status of the four borders:
    # bottom, top, right, left
    def border
      [bottom,top,right,left]
    end
    ##
    # Activate or deactivate all four borders (left, right, top, bottom)
    def border=(boolean)
      [:bottom=, :top=, :right=, :left=].each do |writer| send writer, boolean end
    end
    ##
    # Returns an Array containing the colors of the four borders:
    # bottom, top, right, left
    def border_color
      [@bottom_color,@top_color,@left_color,@right_color]
    end
    ##
    # Set all four border colors to _color_ (left, right, top, bottom)
    def border_color=(color)
      [:bottom_color=, :top_color=, :right_color=, :left_color=].each do |writer|
        send writer, color end
    end
    ##
    # Set the Text rotation
    # Valid values: Integers from -90 to 90,
    # or :stacked (sets #rotation_stacked to true)
    def rotation=(rot)
      if rot.to_s.downcase == 'stacked'
        @rotation_stacked = true
        @rotation = 0
      elsif rot.kind_of?(Integer)
        @rotation_stacked = false
        @rotation = rot % 360
      else
        raise TypeError, "rotation value must be an Integer or the String 'stacked'"
      end
    end
    ##
    # Backward compatibility method. May disappear at some point in the future.
    def center_across!
      self.horizontal_align = :merge
    end
    alias :merge! :center_across!
    ##
    # Is the cell formatted as a Date?
    def date?
      !!Regexp.new(client("[YMD]", 'UTF-8')).match(@number_format.to_s)
    end
    ##
    # Is the cell formatted as a Date or Time?
    def date_or_time?
      !!Regexp.new(client("[hmsYMD]", 'UTF-8')).match(@number_format.to_s)
    end
    ##
    # Is the cell formatted as a DateTime?
    def datetime?
      !!Regexp.new(client("([YMD].*[HS])|([HS].*[YMD])", 'UTF-8')).match(@number_format.to_s)
    end
    ##
    # Is the cell formatted as a Time?
    def time?
      !!Regexp.new(client("[hms]", 'UTF-8')).match(@number_format.to_s)
    end
  end
end
