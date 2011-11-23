# -*- encoding: utf-8 -*-
$:.push File.expand_path("../lib", __FILE__)

Gem::Specification.new do |s|
  s.name        = "nulogy-spreadsheet"
  s.version     = "0.6.5.7.2"
  s.summary     = "The Spreadsheet Library is designed to read and write Spreadsheet Documents"
  s.description = "As of version 0.6.0, only Microsoft Excel compatible spreadsheets are supported"
  s.authors     = ["Masaomi Hatakeyama", "Zeno R.R. Davatz", "Clemens Park"]
  s.email       = ["mhatakeyama@ywesee.com", "zdavatz@ywesee.com", "clemens.park@gmail.com"]
  s.homepage    = "https://github.com/nulogy/spreadsheet"
  s.rubyforge_project = "nulogy-spreadsheet"

  s.files         = `git ls-files`.split("\n")
  s.test_files    = `git ls-files -- {test,spec,features}/*`.split("\n")
  s.executables   = `git ls-files -- bin/*`.split("\n").map{ |f| File.basename(f) }
  s.require_paths = ["lib"]

  s.add_dependency('ruby-ole')
  s.add_development_dependency('rake')
end
