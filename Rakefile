require 'rake'
require 'rake/testtask'
require 'rake/rdoctask'
require 'bundler'
Bundler::GemHelper.install_tasks

desc 'Default: run unit tests.'
task :default => :test

desc 'Test the autocomplete_for plugin.'
Rake::TestTask.new(:test) do |t|
  t.libs << 'lib'
  t.libs << 'test'
  t.pattern = 'test/**/*_test.rb'
  t.verbose = true
end

desc 'Generate documentation for the spreadsheet plugin.'
Rake::RDocTask.new(:rdoc) do |rdoc|
  rdoc.rdoc_dir = 'rdoc'
  rdoc.title    = 'Spreadsheet'
  rdoc.options << '--line-numbers' << '--inline-source'
  rdoc.rdoc_files.include('README.markdown')
  rdoc.rdoc_files.include('lib/**/*.rb')
end
