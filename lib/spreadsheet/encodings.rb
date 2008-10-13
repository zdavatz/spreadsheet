module Spreadsheet
  ##
  # Methods for Encoding-conversions. You should not need to use any of these.
  module Encodings
    if RUBY_VERSION >= '1.9'
      def ascii string
        string.encode 'ASCII//TRANSLIT//IGNORE'
      end
      def client string, internal='UTF-16LE'
        string.encode Spreadsheet.client_encoding
      end
      def internal string, internal='UTF-16LE'
        string.encode internal
      end
    else
      require 'iconv'
      @@utf8_utf16 = Iconv.new('UTF-16LE', 'UTF8')
      @@utf16_ascii = Iconv.new('ASCII//TRANSLIT//IGNORE', 'UTF-16LE')
      @@utf16_utf8 = Iconv.new('UTF8//TRANSLIT//IGNORE', 'UTF-16LE')
      @@iconvs = {}
      def ascii string
        @@utf16_ascii.iconv string
      rescue
        string.gsub /[^\x20-\x7e]+/, ''
      end
      def client string, internal='UTF-16LE'
        key = [Spreadsheet.client_encoding, internal]
        iconv = @@iconvs[key] ||= Iconv.new(Spreadsheet.client_encoding, internal)
        iconv.iconv string
      end
      def internal string, internal='UTF-16LE'
        key = [internal, Spreadsheet.client_encoding]
        iconv = @@iconvs[key] ||= Iconv.new(internal, Spreadsheet.client_encoding)
        iconv.iconv string
      end
    end
  rescue LoadError
    warn "You don't have Iconv support compiled in your Ruby. Spreadsheet may not work as expected"
    def ascii string
      string.gsub /[^\x20-\x7e]+/, ''
    end
    def client string, internal='UTF-16LE'
      string.delete "\0"
    end
    def internal string, internal='UTF-16LE'
      string.split('').zip(Array.new(string.size, 0.chr)).join
    end
  end
end
