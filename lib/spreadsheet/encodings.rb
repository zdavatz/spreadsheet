module Spreadsheet
  ##
  # Methods for Encoding-conversions. You should not need to use any of these.
  module Encodings
    if RUBY_VERSION >= '1.9'
      def client string, internal='UTF-16LE'
        string.force_encoding internal
        string.encode Spreadsheet.client_encoding
      end
      def internal string, client=Spreadsheet.client_encoding
        string.force_encoding client
        string.encode('UTF-16LE').force_encoding('ASCII-8BIT')
      end
      def utf8 string, client=Spreadsheet.client_encoding
        string.force_encoding client
        string.encode('UTF-8')
      end
    else
      require 'iconv'
      @@iconvs = {}
      def client string, internal='UTF-16LE'
        key = [Spreadsheet.client_encoding, internal]
        iconv = @@iconvs[key] ||= Iconv.new(Spreadsheet.client_encoding, internal)
        iconv.iconv string
      end
      def internal string, client=Spreadsheet.client_encoding
        key = ['UTF-16LE', client]
        iconv = @@iconvs[key] ||= Iconv.new('UTF-16LE', client)
        iconv.iconv string
      end
      def utf8 string, client=Spreadsheet.client_encoding
        key = ['UTF-8', client]
        iconv = @@iconvs[key] ||= Iconv.new('UTF-8', client)
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
