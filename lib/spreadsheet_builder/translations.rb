module SpreadsheetBuilder
  TRANSLATIONS = {}
  TRANSLATIONS["text-align"] = Proc.new { |val| 
    allowed = %w{ left center right justify }
    { horizontal_align: val.to_sym } if allowed.include?(val)
  }
  TRANSLATIONS["vertical-align"] = Proc.new { |val|
    allowed = %{ top middle bottom }
    { vertical_align: val.to_sym } if allowed.include?(val)
  }
  TRANSLATIONS["color"] = Proc.new { |val| 
    { color: Palette._color_from_input(val) }
  }
  TRANSLATIONS["background-color"] = Proc.new { |val| 
    { pattern_fg_color: Palette._color_from_input(val), pattern: 1 }
  }
  TRANSLATIONS["font-size"] = Proc.new { |val| 
    { size: SpreadsheetBuilder::CssParser.pt_from_input(val, :height) }
  }
  TRANSLATIONS["font-weight"] = Proc.new { |val| 
    accepted = %{ bold normal }
    if accepted.inlcude?(val)
      { weight: val }
    end
  }
  %w{ border border-top border-bottom border-left border-right border-width border-top-width border-bottom-width border-left-width border-right-width }.each do 
    |key|
    TRANSLATIONS[key] = Proc.new { |val| Border.new(key, val).format }
  end
  # TODO Prove these ratios
  TRANSLATIONS["height"] = Proc.new { |val| 
    { height: SpreadsheetBuilder::CssParser.pt_from_input(val) }  
  }
  TRANSLATIONS["width"] = Proc.new { |val| 
    { width: SpreadsheetBuilder::CssParser.px_from_input(val)  / 7.5 }  
  }
end
