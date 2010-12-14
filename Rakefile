# -*- ruby -*-

$: << File.expand_path("./lib", File.dirname(__FILE__))

require 'rubygems'
require 'hoe'
require './lib/spreadsheet.rb'

ENV['RDOCOPT'] = '-c utf8'

Hoe.spec('spreadsheet') do |p|
   p.developer('Masaomi Hatakeyama, Zeno R.R. Davatz','mhatakeyama@ywesee.com, zdavatz@ywesee.com')
   p.remote_rdoc_dir = ''
   p.extra_deps << ['ruby-ole', '>=1.0']
end

# vim: syntax=Ruby
