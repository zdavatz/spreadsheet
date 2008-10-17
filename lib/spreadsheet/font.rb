require 'spreadsheet/datatypes'
require 'spreadsheet/encodings'

module Spreadsheet
  ##
  # Font formatting data
  class Font
    include Spreadsheet::Datatypes
    include Spreadsheet::Encodings
    attr_accessor :name
    ##
    # You can set the following boolean Font attributes
    # * #italic
    # * #strikeout
    # * #outline
    # * #shadow
    boolean :italic, :strikeout, :outline, :shadow
    ##
    # Font color
    colors :color
    ##
    # Font weight
    # Valid values: :normal, :bold or any positive Integer.
    # In Excel:
    #  100 <= weight <= 1000
    #  :bold   => 700
    #  :normal => 400
    # Default:      :normal
    enum :weight, :normal, :bold, Integer, :bold => :b
    ##
    # Escapement
    # Valid values: :normal, :superscript or :subscript.
    # Default:      :normal
    enum :escapement, :normal, :superscript, :subscript,
         :subscript   => :sub,
         :superscript => :super
    # Font size
    # Valid values: Any positive Integer.
    # Default:      10
    enum :size, 10, Numeric
    # Underline type
    # Valid values: :none, :single, :double, :single_accounting and
    #               :double_accounting.
    # Default:      :none
    enum :underline, :none, :single, :double,
                     :single_accounting, :double_accounting,
         :single => true
    # Font Family
    # Valid values: :none, :roman, :swiss, :modern, :script, :decorative
    # Default:      :none
    enum :family, :none, :roman, :swiss, :modern, :script, :decorative
    # Font Family
    # Valid values: :default, :iso_latin1, :symbol, :apple_roman, :shift_jis,
    #               :korean_hangul, :korean_johab, :chinese_simplified,
    #               :chinese_traditional, :greek, :turkish, :vietnamese,
    #               :hebrew, :arabic, :cyrillic, :thai, :iso_latin2, :oem_latin1
    # Default:      :default
    enum :encoding, :default, :iso_latin1, :symbol, :apple_roman, :shift_jis,
                    :korean_hangul, :korean_johab, :chinese_simplified,
                    :chinese_traditional, :greek, :turkish, :vietnamese,
                    :hebrew, :arabic, :cyrillic, :thai, :iso_latin2, :oem_latin1
    def initialize name, opts={}
      self.name = name
      @color = :text
      opts.each do |key, val|
        self.send "#{key}=", val
      end
    end
    ##
    # Sets #weight to :bold if(_bool_), :normal otherwise.
    def bold= bool
      self.weight = bool ? :bold : nil
    end
    def key # :nodoc:
      key = @name.dup
      key << '_' << size.to_s
      key << '_' << weight.to_s
      key << '_italic'    if italic?
      key << '_strikeout' if strikeout?
      key << '_outline'   if outline?
      key << '_shadow'    if shadow?
      key << '_' << escapement.to_s
      key << '_' << underline.to_s
      key << '_' << color.to_s
      key << '_' << family.to_s
      key << '_' << encoding.to_s
    end
  end
end
