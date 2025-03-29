# -*- ruby -*-
require "bundler/gem_tasks"
require "standard/rake"

desc "test using minittest via test/suite.rb"
task :test do |t|
  $LOAD_PATH << File.dirname(__FILE__)
  require "test/suite"
end
# vim: syntax=Ruby
