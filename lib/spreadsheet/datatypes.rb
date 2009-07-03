require 'spreadsheet/compatibility'

module Spreadsheet
  ##
  # This module defines convenience-methods for the definition of Spreadsheet
  # attributes (boolean, colors and enumerations)
  module Datatypes
    include Compatibility
    def Datatypes.append_features mod
      super
      mod.module_eval do
class << self
  ##
  # Valid colors for color attributes.
  COLORS = [ :builtin_black, :builtin_white, :builtin_red, :builtin_green,
             :builtin_blue, :builtin_yellow, :builtin_magenta, :builtin_cyan,
             :text, :border, :pattern_bg, :dialog_bg, :chart_text, :chart_bg,
             :chart_border, :tooltip_bg, :tooltip_text, :aqua,
             :black, :blue, :cyan, :brown, :fuchsia, :gray, :grey, :green,
             :lime, :magenta, :navy, :orange, :purple, :red, :silver, :white,
             :yellow ]
  ##
  # Define instance methods to read and write boolean attributes.
  def boolean *args
    args.each do |key|
      define_method key do
        name = ivar_name key
        !!(instance_variable_get(name) if instance_variables.include?(name))
      end
      define_method "#{key}?" do
        send key
      end
      define_method "#{key}=" do |arg|
        arg = false if arg == 0
        instance_variable_set(ivar_name(key), !!arg)
      end
      define_method "#{key}!" do
        send "#{key}=", true
      end
    end
  end
  ##
  # Define instance methods to read and write color attributes.
  # For valid colors see COLORS
  def colors *args
    args.each do |key|
      attr_reader key
      define_method "#{key}=" do |name|
        name = name.to_s.downcase.to_sym
        if COLORS.include?(name)
          instance_variable_set ivar_name(key), name
        else
          raise ArgumentError, "unknown color '#{name}'"
        end
      end
    end
  end
  ##
  # Define instance methods to read and write enumeration attributes.
  # * The first argument designates the attribute name.
  # * The second argument designates the default value.
  # * All subsequent attributes are possible values.
  # * If the last attribute is a Hash, each value in the Hash designates
  #   aliases for the corresponding key.
  def enum key, *values
    aliases = {}
    if values.last.is_a? Hash
      values.pop.each do |value, synonyms|
        if synonyms.is_a? Array
          synonyms.each do |synonym| aliases.store synonym, value end
        else
          aliases.store synonyms, value
        end
      end
    end
    values.each do |value|
      aliases.store value, value
    end
    define_method key do
      name = ivar_name key
      value = instance_variable_get(name) if instance_variables.include? name
      value || values.first
    end
    define_method "#{key}=" do |arg|
      if arg
        arg = aliases.fetch arg do
          aliases.fetch arg.to_s.downcase.gsub(/[ \-]/, '_').to_sym, arg
        end
        if values.any? do |val| val === arg end
          instance_variable_set(ivar_name(key), arg)
        else
          valid = values.collect do |val| val.inspect end.join ', '
          raise ArgumentError,
            "Invalid value '#{arg.inspect}' for #{key}. Valid values are: #{valid}"
        end
      else
        instance_variable_set ivar_name(key), values.first
      end
    end
  end
end
      end
    end
  end
end
