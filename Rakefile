# -*- ruby -*-

$: << File.expand_path("./lib", File.dirname(__FILE__))

require 'rubygems'
require 'hoe'
require './lib/spreadsheet.rb'

ENV['RDOCOPT'] = '-c utf8'

Hoe.new('spreadsheet', Spreadsheet::VERSION) do |p|
  # p.rubyforge_name = 'spreadsheetx' # if different than lowercase project name
   p.developer('Masaomi Hatakeyama, Zeno R.R. Davatz','mhatakeyama@ywesee.com, zdavatz@ywesee.com')
   p.remote_rdoc_dir = ''
   p.extra_deps << 'ruby-ole'
end

# vim: syntax=Ruby
