module Spreadsheet
  ##
  # The Row class. Encapsulates Cell data and formatting.
  # Since Row is a subclass of Array, you may use all the standard Array methods
  # to manipulate a Row.
  # By convention, Row#at will give you raw values, while Row#[] may be
  # overridden to return enriched data if necessary (see also the Date- and
  # DateTime-handling in Excel::Row#[]
  #
  # Useful Attributes are:
  # #idx::            The 0-based index of this Row in its Worksheet.
  # #formats::        A parallel array containing Formatting information for
  #                   all cells stored in a Row.
  # #default_format:: The default Format used when writing a Cell if no explicit
  #                   Format is stored in #formats for the cell.
  # #height::         The height of this Row in points (defaults to 12).
  class Row < Array
    class << self
      def updater *keys
        keys.each do |key|
          define_method key do |*args|
            res = super
            @worksheet.row_updated @idx, self if @worksheet
            res
          end
        end
      end
    end
    attr_reader :formats, :default_format
    attr_accessor :idx, :height, :worksheet
    updater :[]=, :clear, :concat, :delete, :delete_if, :fill, :insert, :map!,
            :pop, :push, :reject!, :replace, :reverse!, :shift, :slice!,
            :sort!, :uniq!, :unshift
    def initialize worksheet, idx, cells=[]
      @worksheet = worksheet
      @idx = idx
      while !cells.empty? && !cells.last
        cells.pop
      end
      super cells
      @first_used ||= index_of_first self
      @first_unused ||= size
      @formats = []
      @height = 12
    end
    ##
    # Set the default Format used when writing a Cell if no explicit Format is
    # stored for the cell.
    def default_format= format
      @worksheet.add_format format if @worksheet
      @default_format = format
    end
    ##
    # #first_unused (really last used + 1) - the 0-based index of the first of
    # all remaining contiguous blank Cells.
    alias :first_unused :size
    ##
    # #first_used the 0-based index of the first non-blank Cell.
    def first_used
      index_of_first self
    end
    ##
    # The Format for the Cell at _idx_ (0-based), or the first valid Format in
    # Row#default_format, Column#default_format and Worksheet#default_format.
    def format idx
      @formats[idx] || @default_format \
        || @worksheet.column(idx).default_format if @worksheet
    end
    ##
    # Set the Format for the Cell at _idx_ (0-based).
    def set_format idx, fmt
      @formats[idx] = fmt
      @worksheet.add_format fmt
      @worksheet.row_updated @idx, self if @worksheet
      fmt
    end
    def inspect
      variables = instance_variables.collect do |name|
        "%s=%s" % [name, instance_variable_get(name)]
      end.join(' ')
      sprintf "#<%s:0x%014x %s %s>", self.class, object_id, variables, super
    end
    private
    def index_of_first ary # :nodoc:
      ary.index(ary.find do |elm| elm end)
    end
  end
end
