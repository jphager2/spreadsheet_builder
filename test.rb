require 'css_parser'
include CssParser

parser = CssParser::Parser.new
parser.load_uri("file://#{Dir.pwd}/test.css")

accepted_keys = %w{ color border-color font-size font-weight }

parser.each_rule_set do |rset|
  rset.each_declaration do |key,_|
    rset.remove_declaration!(key) unless accepted_keys.include?(key)
  end
end
