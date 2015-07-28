require 'shade'
require 'nokogiri'
require 'css_parser'
require 'spreadsheet'

require_relative 'spreadsheet_builder/border'
require_relative 'spreadsheet_builder/translations'

# TODO find out if this is necessary
include CssParser

class SpreadsheetBuilder

  # Utility
  def self.merge(*options)
    options.inject(&:merge)
  end
end
