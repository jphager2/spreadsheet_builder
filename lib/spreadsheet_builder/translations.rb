module SpreadsheetBuilder
  TRANSLATIONS = {}
  TRANSLATIONS["text-align"] = Proc.new { |val| 
    allowed = %w{ left center right justify }

    if allowed.include?(val)
      { horizontal_align: val.to_sym }
    end
  }
  TRANSLATIONS["color"] = Proc.new { |val| 
    { color: Palette._color_from_input(val) }
  }
  TRANSLATIONS["background-color"] = Proc.new { |val| 
    { pattern_fg_color: Palette._color_from_input(val), pattern: 1 }
  }
  TRANSLATIONS["font-size"] = Proc.new { |val| 
    { size: Integer(val) }
  }
  TRANSLATIONS["font-weight"] = Proc.new { |val| 
    accepted = %{ bold normal }
    if accepted.inlcude?(val.to_s)
      { weight: val.to_s }
    end
  }
  %w{ border border-top border-bottom border-left border-right border-width border-top-width border-bottom-width border-left-width border-right-width }.each do 
    |key|
    TRANSLATIONS[key] = Proc.new { |val| Border.new(key, val).format }
  end
  TRANSLATIONS["height"] = Proc.new { |val| { height: val.to_i / 2.0 }  }
  TRANSLATIONS["width"] = Proc.new { |val| { width: val.to_i / 7.5 }  }
end
