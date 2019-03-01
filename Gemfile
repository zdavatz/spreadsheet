source "https://rubygems.org"

if ENV['USE_LATEST_RUBY_OLE']
  if Dir.exist?('../ruby-ole')
    gem 'ruby-ole', :path => '../ruby-ole'
  else
    gem 'ruby-ole',
      :git => 'https://github.com/taichi-ishitani/ruby-ole.git',
      :branch => 'support_frozen_string_literal'
  end
else
  gem 'ruby-ole'
end

if RUBY_VERSION.to_f > 2.0
  gem 'test-unit'
  gem 'minitest'
end
group :development do
  gem 'hoe', '>= 3.4'
end
