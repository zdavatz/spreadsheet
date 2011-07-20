require "rubygems"
require "rake"

spec = Gem::Specification.new do |s|
   s.name        = "spreadsheet"
   s.version     = "0.6.5.7"
   s.summary     = "The Spreadsheet Library is designed to read and write Spreadsheet Documents"
   s.description = "As of version 0.6.0, only Microsoft Excel compatible spreadsheets are supported"
   s.author      = "Masaomi Hatakeyama, Zeno R.R. Davatz"
   s.email       = "mhatakeyama@ywesee.com, zdavatz@ywesee.com"
   s.platform    = Gem::Platform::RUBY
   s.files       = Dir.glob("{bin,lib,test}/**/*") + Dir.glob("*.txt")
   s.test_file   = "test/suite.rb"
   s.executables << 'xlsopcodes'
   s.add_dependency('ruby-ole')
   s.homepage	 = "http://scm.ywesee.com/?p=spreadsheet/.git;a=summary"
end

if $0 == __FILE__
   Gem.manage_gems
   Gem::Builder.new(spec).build
end
