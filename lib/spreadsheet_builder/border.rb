class Border

  def initialize(key, val, dir = nil)
    @dir     = dir && dir.to_sym
    @vals    = val.scan(/[^\s]+/)
    @keys    = key.scan(/[^-]+/)
    @border  = self
    dir_hash = { top: nil, bottom: nil, left: nil, right: nil }
    @css     = { 
      color: dir_hash.dup,
      style: dir_hash.dup,
      width: dir_hash.dup
    }

    case @keys[1]
    when Proc.new { |type| @css.keys.include?(type.to_sym) if type }
      assign_vals_to(@keys[1])
    when nil # "border" 
      width, style, color = @vals
      w, s, c = @css[:width], @css[:style], @css[:color]

      w[:top] = w[:bottom] = w[:left] = w[:right] = width
      s[:top] = s[:bottom] = s[:left] = s[:right] = style
      c[:top] = c[:bottom] = c[:left] = c[:right] = color
    else     # "border-left"
      dir     = @keys.delete_at(1)
      @border = Border.new(@keys.join('-'), @vals.join(' '), dir)
    end
  end

  #border: 1px solid black;
  #border-width: 1px 1px 1px 1px;
  #border-width: 1px; // all
  #border-width: 1px 2px; // top and bottom 1px, left and right 2px
  #border-width: 1px 2px 3px; // top 1px, left and right 2px, bottom 3px
  #border-#{direction}-width: 1px;

  def assign_vals_to(type)
    t = @css[type.to_sym]
    case @vals.length
    when 1
      t[:top] = t[:bottom] = t[:left] = t[:right] = @vals[0]
    when 2
      t[:top]  = t[:bottom] = @vals[0]
      t[:left] = t[:right]  = @vals[1]
    when 3
      t[:top]              = @vals[0]
      t[:left] = t[:right] = @vals[1]
      t[:bottom]           = @vals[2]
    when 4
      t[:top], t[:bottom], t[:left], t[:right] = @vals
    end
  end

  def format
    @border._format
  end

  SPREADSHEET_VALUES = [
    :none,
    :thin, 
    :medium, 
    :thick, 
    :double, 
    :hair, 
    :dashed, 
    :dotted, 
    :thin_dash_dotted,
    :thin_dash_dot_dotted,
    :medium_dashed, 
    :medium_dash_dotted, 
    :medium_dash_dot_dotted, 
    :slanted_medium_dash_dotted 
  ]

  colors_attrs = [
    :bottom_color, 
    :top_color, 
    :left_color, 
    :right_color,
    :pattern_fg_color, 
    :pattern_bg_color,
    :diagonal_color
  ]

  protected
  def _format
    format = {}

    # merge width and style
    translated_values.values.first.keys.each do  |dir|
      width = translated_values[:width][dir] 
      style = translated_values[:style][dir]

      if width && style && width != style
        found = SPREADSHEET_VALUES.find { |val|
          val.to_s =~ /#{width}/ && val.to_s =~ /#{style}/
        }
        if found
          translated_values[:width][dir] = found
          translated_values[:style][dir] = found
        end
      end
    end

    dirs = @dir ? [@dir] : translated_values.values.first.keys
    dirs.each do |dir|
      style = translated_values[:width][dir] || translated_values[:style][dir]
      color = translated_values[:color][dir]

      format["#{dir}_color".to_sym] = color if color
      format[dir.to_sym]            = style if style
      format[:pattern_fg_color]     = color if style.to_s =~ /(dash|dott)/
    end
    format
  end

  private
  def translated_values
    return @translated_values if @translated_values

    @translated_values = {}
    @css.each { |k,v| @translated_values[k] = v.dup }

    @translated_values.each do |type, vals|
      vals.each do |dir, v|
        @translated_values[type][dir] = __send__("translate_#{type}", v)
      end
    end
  end

  #thin|medium|thick
  def translate_width(val)
    return unless val

    if val.to_i.to_s == val.to_s
      val = val.to_i
      if    val <= 0 then nil
      elsif val 1    then :thin
      elsif val 2    then :medium
      else                :thick
      end
    elsif SPREADSHEET_VALUES.include?(val.to_sym)
      val.to_sym
    end
  end

  def translate_color(val)
    return unless val

    SpreadsheetBuilder::Palette._color_from_input(val)
  end

  def translate_style(val)
    return unless val

    if SPREADSHEET_VALUES.include?(val.to_sym) 
      val.to_sym
    else
      :thin
    end
  end
end
