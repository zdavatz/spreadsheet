require 'date'
require 'spreadsheet/column'
require 'spreadsheet/encodings'
require 'spreadsheet/row'

module Spreadsheet
  ##
  # The Worksheet class. Contains most of the Spreadsheet data in Rows.
  #
  # Interesting Attributes
  # #name          :: The Name of this Worksheet.
  # #default_format:: The default format used for all cells in this Workhseet
  #                   that have no format set explicitly or in
  #                   Row#default_format.
  # #rows          :: The Rows in this Worksheet. It is not recommended to
  #                   Manipulate this Array directly. If you do, call
  #                   #updated_from with the smallest modified index.
  # #columns       :: The Column formatting in this Worksheet. Column
  #                   instances may appear at more than one position in #columns.
  #                   If you modify a Column directly, your changes will be
  #                   reflected in all those positions.
  class Worksheet
    include Spreadsheet::Encodings
    include Enumerable
    attr_accessor :name, :workbook
    attr_reader :rows, :columns
    def initialize opts={}
      @dimensions = [0,0,0,0]
      @name = opts[:name] || 'Worksheet'
      @workbook = opts[:workbook]
      @rows = []
      @columns = []
      @links = {}
    end
    ##
    # Add a Format to the Workbook. If you use Row#set_format, you should not
    # need to use this Method.
    def add_format fmt
      @workbook.add_format fmt if fmt
    end
    ##
    # Get the enriched value of the Cell at _row_, _column_.
    # See also Worksheet#[], Row#[].
    def cell row, column
      row(row)[column]
    end
    ##
    # Returns the Column at _idx_.
    def column idx
      @columns[idx] || Column.new(idx, default_format, :worksheet => self)
    end
    ##
    # The number of columns in this Worksheet which contain data.
    def column_count
      dimensions[3] - dimensions[2]
    end
    def column_updated idx, column
      @columns[idx] = column
    end
    ##
    # Delete the Row at _idx_ (0-based) from this Worksheet.
    def delete_row idx
      res = @rows.delete_at idx
      updated_from idx
      res
    end
    ##
    # The default Format of this Worksheet, if you have set one.
    # Returns the Workbook's default Format otherwise.
    def default_format
      @default_format || @workbook.default_format
    end
    ##
    # Set the default Format of this Worksheet.
    def default_format= format
      @default_format = format
      add_format format
      format
    end
    ##
    # Dimensions:: [ first used row, first unused row,
    #              first used column, first unused column ]
    #              ( First used means that all rows or columns before that are
    #              empty. First unused means that this and all following rows
    #              or columns are empty. )
    def dimensions
      @dimensions || recalculate_dimensions
    end
    ##
    # If no argument is given, #each iterates over all used Rows (from the first
    # used Row until but omitting the first unused Row, see also #dimensions).
    #
    # If the argument skip is given, #each iterates from that row until but
    # omitting the first unused Row, effectively skipping the first _skip_ Rows
    # from the top of the Worksheet.
    def each skip=dimensions[0], &block
      skip.upto(dimensions[1] - 1) do |idx|
        block.call row(idx)
      end
    end
    def encoding # :nodoc:
      @workbook.encoding
    end
    ##
    # Sets the default Format of the column at _idx_.
    #
    # _idx_ may be an Integer, or an Enumerable that iterates over a number of
    # Integers.
    #
    # _format_ is a Format, or nil if you want to remove the Formatting at _idx_
    #
    # Returns an instance of Column if _idx_ is an Integer, an Array of Columns
    # otherwise.
    def format_column idx, format=nil, opts={}
      opts[:worksheet] = self
      res = case idx
            when Integer
              column = nil
              if format
                column = Column.new(idx, format, opts)
              end
              @columns[idx] = column
            else
              idx.collect do |col| format_column col, format, opts end
            end
      shorten @columns
      res
    end
    ##
    # Insert a Row at _idx_ (0-based) containing _cells_
    def insert_row idx, cells=[]
      res = @rows.insert idx, Row.new(self, idx, cells)
      updated_from idx
      res
    end
    def inspect
      "#<#{self.class}:#{object_id} "
      names = instance_variables
      names.delete '@rows'
      variables = names.collect do |name|
        "%s=%s" % [name, instance_variable_get(name)]
      end.join(' ')
      sprintf "#<%s:0x%014x %s @rows[%i]>", self.class, object_id,
                                            variables, row_count
    end
    ## The last Row containing any data
    def last_row
      row(last_row_index)
    end
    ## The index of the last Row containing any data
    def last_row_index
      [dimensions[1] - 1, 0].max
    end
    ##
    # Replace the Row at _idx_ with the following arguments. Like #update_row,
    # but truncates the Row if there are fewer arguments than Cells in the Row.
    def replace_row idx, *cells
      if(row = @rows[idx]) && cells.size < row.size
        cells.concat Array.new(row.size - cells.size)
      end
      update_row idx, *cells
    end
    ##
    # The Row at _idx_ or a new Row.
    def row idx
      @rows[idx] || Row.new(self, idx)
    end
    ##
    # The number of Rows in this Worksheet which contain data.
    def row_count
      dimensions[1] - dimensions[0]
    end
    ##
    # Tell Worksheet that the Row at _idx_ has been updated and the #dimensions
    # need to be recalculated. You should not need to call this directly.
    def row_updated idx, row
      @dimensions = nil
      row = @rows[idx] = shorten(row)
      format_dates row
      row
    end
    ##
    # Updates the Row at _idx_ with the following arguments.
    def update_row idx, *cells
      res = if row = @rows[idx]
              row[0, cells.size] = cells
              row
            elsif cells = shorten(cells)
              Row.new self, idx, cells
            end
      row_updated idx, res
      res
    end
    ##
    # Renumbers all Rows starting at _idx_ and calls #row_updated for each of
    # them.
    def updated_from idx
      idx.upto(@rows.size - 1) do |idx|
        row = row(idx)
        row.idx = idx
        row_updated idx, row
      end
    end
    ##
    # Get the enriched value of the Cell at _row_, _column_.
    # See also Worksheet#cell, Row#[].
    def [] row, column
      row(row)[column]
    end
    ##
    # Set the value of the Cell at _row_, _column_ to _value_.
    # See also Row#[]=.
    def []= row, column, value
      row(row)[column] = value
    end
    private
    def format_dates row # :nodoc:
      ## If no format is set in a cell which contains a Date or Time, we will
      #  add a format. At the moment, the number formats correspond to builtin
      #  Excel number-formats. We may need to add a level of abstraction if
      #  additional writers are added to the library.
      return unless row
      row.each_with_index do |value, idx|
        unless row.formats[idx]
          format = nil
          case value
          when Date
            format = @workbook.formats.find do |fmt| fmt.date? end
            format ||= Format.new :number_format => client('M/D/YY', 'UTF-8')
          when DateTime, Time
            format = @workbook.formats.find do |fmt| fmt.datetime? end
            format ||= Format.new :number_format => client('M/D/YY h:mm', 'UTF-8')
          end
          if format
            row.formats[idx] = format
            @workbook.add_format format
          end
        end
      end
    end
    def index_of_first ary # :nodoc:
      return unless ary
      ary.index(ary.find do |elm| elm end)
    end
    def recalculate_dimensions # :nodoc:
      shorten @rows
      @dimensions = []
      @dimensions[0] = index_of_first(@rows) || 0
      @dimensions[1] = @rows.size
      compact = @rows.compact
      @dimensions[2] = compact.collect do |row| index_of_first row end.compact.min || 0
      @dimensions[3] = compact.collect do |row| row.size end.max || 0
      @dimensions
    end
    def shorten ary # :nodoc:
      return unless ary
      while !ary.empty? && !ary.last
        ary.pop
      end
      ary unless ary.empty?
    end
  end
end
