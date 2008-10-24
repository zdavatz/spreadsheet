module Spreadsheet
  ##
  # Methods for Encoding-conversions. You should not need to use any of these.
  module Encodings
    if RUBY_VERSION >= '1.9'
      def client string, internal='UTF-16LE'
        string.encode Spreadsheet.client_encoding
      end
      def internal string, internal='UTF-16LE'
        string.encode internal
      end
    else
      require 'iconv'
      @@iconvs = {}
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
    def client string, internal='UTF-16LE'
      string.delete "\0"
    end
    def internal string, internal='UTF-16LE'
      string.split('').zip(Array.new(string.size, 0.chr)).join
    end
  end
end
