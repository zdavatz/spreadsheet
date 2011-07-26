#!/usr/bin/env ruby
# suite.rb -- spreadsheet -- 26.07.2011 -- zdavatz@ywesee.com

require 'find'

here = File.dirname(__FILE__)

$: << here

Find.find(here) do |file|
	if /(?<!suite)\.rb$/o.match(file)
    #from Roel van der Hoorn vanderhoorn@gmail.com
    #should work for Ruby 1.8 and 1.9, without Oniguruma
#  if /(?:^|\/)(?!suite)[^\/]+\.rb$/o
    require file
	end
end
