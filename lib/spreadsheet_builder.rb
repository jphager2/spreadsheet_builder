require 'shade'
require 'nokogiri'
require 'css_parser'
require 'spreadsheet'

require_relative 'spreadsheet_builder/builder'
require_relative 'spreadsheet_builder/css_parser'
require_relative 'spreadsheet_builder/css_rule'
require_relative 'spreadsheet_builder/html_parser'
require_relative 'spreadsheet_builder/border'
require_relative 'spreadsheet_builder/data'
require_relative 'spreadsheet_builder/palette'
require_relative 'spreadsheet_builder/translations'

# TODO find out if this is necessary
include CssParser

module SpreadsheetBuilder

  # Utility
  def self.merge(*options)
    options.inject(&:merge)
  end
end
