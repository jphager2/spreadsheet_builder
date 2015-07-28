module SpreadsheetBuilder
  Rgb = Spreadsheet::Excel::Rgb

  CUSTOM_PALETTE = {
    :xls_color_0 => Rgb.new(0,0,0),
    :xls_color_1 => Rgb.new(255,255,255),
    :xls_color_2 => Rgb.new(204,204,204),
    :xls_color_3 => Rgb.new(249,249,249)
  }

  PALETTE = Shade::Palette.new do |p| 
    Rgb.class_variable_get(:@@RGB_MAP).merge(CUSTOM_PALETTE).each do 
      |name, value|
      p.add("##{value.to_i.to_s(16).ljust(6, "0")}", name.to_s)
    end
  end

  module Palette
    def self._color_from_input(input)
      input = input.to_s
      if input =~ /^rgb/i
        _, r, g, b = input.match(/^rgba*\(\s*(\d+)\s*,\s*(\d+)\s*,\s*(\d+)[^\)]*/)
        r, g, b = [r, g, b].map(&:to_i)
        input = "##{Spreadsheet::Excel::Rgb.new(r, g, b).as_hex.ljust(6, "0")}"
      end

      # Assume a color is always found
      color = PALETTE.nearest_value(input).name.to_sym
      color
    end
  end
end
