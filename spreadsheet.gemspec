require File.join(File.dirname(__FILE__), 'lib', 'spreadsheet')

spec = Gem::Specification.new do |s|
   s.name        = "spreadsheet"
   s.version     =  Spreadsheet::VERSION
   s.summary     = "The Spreadsheet Library is designed to read and write Spreadsheet Documents"
   s.description = "As of version 0.6.0, only Microsoft Excel compatible spreadsheets are supported"
   s.author      = "Masaomi Hatakeyama, Zeno R.R. Davatz"
   s.email       = "mhatakeyama@ywesee.com, zdavatz@ywesee.com"
   s.platform    = Gem::Platform::RUBY
   s.license     = "GPLv3"
   s.files       = Dir.glob("{bin,lib,test}/**/*") + Dir.glob("*.txt")
   s.test_file   = "test/suite.rb"
   s.executables << "xlsopcodes"

   s.add_dependency "ruby-ole"
   s.add_development_dependency "hoe"

   s.homepage	 = "https://github.com/zdavatz/spreadsheet/"
end

