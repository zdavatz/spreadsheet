require_relative 'lib/spreadsheet/version'

Gem::Specification.new do |spec|
   spec.name        = "spreadsheet"
   spec.version     =  Spreadsheet::VERSION
   spec.homepage    = "https://github.com/zdavatz/spreadsheet"
   spec.summary     = "The Spreadsheet Library is designed to read and write Spreadsheet Documents"
   spec.description = "As of version 0.6.0, only Microsoft Excel compatible spreadsheets are supported"
   spec.author      = "Hannes F. Wyss, Masaomi Hatakeyama, Zeno R.R. Davatz"
   spec.email       = "hannes.wyss@gmail.com, mhatakeyama@ywesee.com, zdavatz@ywesee.com"
   spec.platform    = Gem::Platform::RUBY
   spec.license     = "GPL-3.0"
   spec.files       = Dir.glob("{bin,lib,test}/**/*") + Dir.glob("*.txt")
   spec.test_file   = "test/suite.rb"
   spec.executables << "xlsopcodes"

   spec.add_dependency "bigdecimal"
   spec.add_dependency "ruby-ole"
   spec.add_development_dependency "rake"
   spec.add_development_dependency "test-unit"
   spec.add_development_dependency "simplecov"

   spec.homepage    = 'https://github.com/zdavatz/spreadsheet'
   spec.metadata["changelog_uri"] = spec.homepage + "/blob/master/History.md"
   spec.metadata["funding_uri"] = "https://github.com/sponsors/zdavatz"
end
