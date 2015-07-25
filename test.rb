require 'css_parser'
include CssParser

parser = CssParser::Parser.new
parser.load_uri!("file://#{Dir.pwd}/test.css")

accepted_keys = %w{ color border-color font-size font-weight text-align }

parser.each_rule_set do |rset|
  rset.each_declaration do |key,_|
    rset.remove_declaration!(key) unless accepted_keys.include?(key)
  end
end

Translations = {}
Translations["text-align"] = Proc.new { |val| 
  allowed = %w{ left center right justify }
  if allowed.include?(val)
    { horizontal_align: val.to_sym }
  end
}

def translate_declaration(dec)
  #here is the real work
  #puts "dec: #{dec}"
  key, val = dec.sub(/;$/,'').gsub(/\s/, '').split(':')
  #puts "key: #{key}; val: #{val}"
  if key && val
    format = Translations[key].call(val)
  end
  format || {}
end

def format_for(declarations)
  #puts "declarations: #{declarations}"
  declarations.delete_if { |dec| dec && dec.empty? }
  declarations.each_with_object({}) { |dec, format|
    format.merge!(translate_declaration(dec))
  }
end

selectors = %w{ .text-center .text-left .text-justify .text-right }

selectors.each do |selector|
  puts "looking for #{selector}"
  if found = parser.find_by_selector(selector)
    puts format_for(found).inspect
  else
    puts "not found"
  end
end
