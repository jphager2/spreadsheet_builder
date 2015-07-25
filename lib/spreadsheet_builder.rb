require 'nokogiri'
require 'spreadsheet'
require 'shade'
require 'css_parser'
include CssParser
require_relative 'border'

class SpreadsheetBuilder

  # Examples
  #H1 = { weight: :bold }
  #SOMETHING_ELSE = { color: :red }
  #TEXT_CENTER = { horizontal_align: :center }
  #

  PALETTE = Shade::Palette.new do |p| 
    Spreadsheet::Excel::Rgb.class_variable_get(:@@RGB_MAP).each do |name, value|
      p.add("##{value.to_s(16).ljust(6, "0")}", name.to_s)
    end
  end

  # TODO document this
  def self.from_slim(file, options = {}, context = self, &block)
    html = Slim::Template.new(file, options).render(context, &block)
    from_html(html)
  end

  def self.from_erb(file)
    html     = File.read(file)
    template = ERB.new(html)
    html     = template.result

    from_html(html)
  end

  def self.merge(*options)
    options.inject(&:merge)
  end

  TRANSLATIONS = {}
  TRANSLATIONS["text-align"] = Proc.new { |val| 
    allowed = %w{ left center right justify }

    if allowed.include?(val)
      { horizontal_align: val.to_sym }
    end
  }
  TRANSLATIONS["color"] = Proc.new { |val| 
    { color: _color_from_input(val) }
  }
  TRANSLATIONS["background-color"] = Proc.new { |val| 
    { pattern_fg_color: _color_from_input(val), pattern: 1 }
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

  def self._translate_declaration(key, val)
    f = TRANSLATIONS[key].call(val) if key && val
    f || {}
  end

  def self._format_for_declarations(declarations)
    declarations.delete_if { |_,v| v.nil? || v.empty? }

    declarations.each_with_object({}) { |(k,v), format|
      format.merge!(_translate_declaration(k,v[:value]))
    }
  end

  def self._format_from_klass_tree(klass)
    if _css_cache[klass]
      format = _css_cache[klass]
    else
      rules = _find_rules(klass).sort_by { |specificity,_| specificity }
      declarations = rules.map { |_,rules| 
        rules.inject({}) { |dec, r| 
          dec.merge(r.instance_variable_get(:@declarations))
        }
      }.inject({},&:merge)
      format = _format_for_declarations(declarations)
      _css_cache[klass] = format
    end

    format || {}
  end

  def self._css_rules
    @css_rules || _reset_css_rules
  end

  def self._css_cache
    @css_cache || _reset_css_cache
  end

  def self._css_parser
    @css_parser || _reset_css_parser
  end

  def self._reset_css_rules(parser = _css_parser)
    @css_rules = []
    parser.each_rule_set(:all) { |r| @css_rules << r }
    @css_rules
  end

  def self._reset_css_cache
    @css_cache = {}
  end
  

  def self._reset_css_parser
    parser = CssParser::Parser.new
    # TODO load these files from a config
    # parser.load_uri!("file://#{Dir.pwd}/test.css")
    # TODO or even better parse the html doc for spreadsheet links 
    # and load those
    parser.load_uri!("file://#{Dir.pwd}/test2.css")

    accepted_keys = %w{ color background-color font-size font-weight text-align border border-width border-style border-color }
    dirs = %w{ top bottom left right }
    types = %w{ width style color }
    dirs.each do |dir|
      accepted_keys << "border-#{dir}"
      types.each do |type|
        accepted_keys << "border-#{dir}-#{type}"
      end
    end

    # Explicity reset rules to avoid infinite loop or bad data
    _reset_css_rules(parser)
    _css_rules.each do |rule|
      rule.each_declaration do |key,_|
        rule.remove_declaration!(key) unless accepted_keys.include?(key)
      end
    end
     _css_rules.delete_if { |rule| 
       rule.instance_variable_get(:@declarations).empty?
     }

    _reset_css_cache
    @css_parser = parser
  end

  def self._add_td_to_cells(row, col, td, tr_format, cells)
    found = cells.find { |cell| cell[:row] == row && cell[:col] == col}
    unless found 
      td_klass  = _klass_tree_for_node(td)
      td_format = tr_format.merge(_format_from_klass_tree(td_klass))
      cells << { 
        row: row, 
        col: col, 
        value: td.text.strip, 
        format: td_format, 
        path: td.css_path 
      } 
    else
      _add_td_to_cells(row, col + 1, td, tr_format, cells)
    end
  end

  # TODO support :first-child, :nth-child(n|even|odd), :last-child, :first-of-kind, :last-of-kind, ... ?
  def self._klass_node_for_node(node)
    name     = node.name
    root     = node.ancestors.last
    siblings = node.parent.element_children

    index         = siblings.index(node) + 1
    first_of_kind = root.css(name).first == node
    last_of_kind  = root.css(name).last  == node

    klasses = node.attributes["class"] && node.attributes["class"].value
    klasses = klasses.to_s.split(/ /).map { |k| "." + k }
    klasses << name
    klasses << "#{name}:nth-child(#{index})"
    klasses << "#{name}:nth-child(odd)"  if index.odd?
    klasses << "#{name}:nth-child(even)" if index.even?
    klasses << "#{name}:first-child)"    if index == 1
    klasses << "#{name}:last-child)"     if index == siblings.length
    klasses << "#{name}:first-of-kind)"  if first_of_kind
    klasses << "#{name}:last-of-kind)"   if last_of_kind
    klasses 
  end

  def self._klass_tree_for_node(base)
    tree = (base.ancestors.reverse.drop(1) << base)
    tree.map { |n, t| _klass_node_for_node(n) }
  end

  def self._find_rules(klass_tree)
    rules = Hash.new { |h,k| h[k] = [] } 
    _css_rules.each do |rule|
      selectors = rule.selectors.map { |s| 
        s = s.split(/[\s>]/).map { |node| node.split('.') }
        s.each do |node| 
          n = node.length
          node[1...n] = node[1...n].map { |k| '.' + k }
          node.delete_if { |k| k.empty? || k == "." }
        end 
      }
      found = selectors.sort_by(&:length).find { |s| 
        next unless klass_tree.last & s.last == s.last

        s[0..-2].reverse.inject(klass_tree.length - 2) do |i, s_node|
          break false unless i
          klass_tree[0..i].index { |kt_node| kt_node & s_node == s_node }
        end
      }
      rules[found.length] << rule if found
    end
    rules
  end

  def self.from_html(html, force_level = :none)
    _reset_css_parser if force_level == :parser
    _reset_css_cache  if force_level == :cache
    _reset_css_rules  if force_level == :rules

    # need to check for attributes colspan and row span
    cells       = [] 
    merges      = []
    col_widths  = {}
    row_heights = {}

    doc      = Nokogiri::HTML(html)
    tb       = doc.css('table').first 

    # ignoring specified formats for anything other than table tr td/th
    tb_klass  = _klass_tree_for_node(tb)
    tb_format = _format_from_klass_tree(tb_klass) 

    doc.css('tr').each_with_index do |tr, row|
      tr_klass  = _klass_tree_for_node(tr)
      tr_format = tb_format.merge(_format_from_klass_tree(tr_klass))

      tr.css('td, th').each_with_index do |td, col|
         
        rowheight = td.attributes["rowheight"]
        colwidth  = td.attributes["colwidth"]
        rowspan   = td.attributes["rowspan"]
        colspan   = td.attributes["colspan"]

        rowheight &&= rowheight.value.to_i
        colwidth  &&= colwidth.value.to_i
        rowspan   &&= rowspan.value.to_i
        colspan   &&= colspan.value.to_i

        _add_td_to_cells(row, col, td, tr_format, cells)
        if colspan
          (1..colspan-1).each {|t| 
            _add_td_to_cells(row, col+t, td, tr_format, cells)
          }
        end
        if rowspan
          (1..rowspan-1).each {|t| 
            _add_td_to_cells(row+t, col, td, tr_format, cells)
          }
        end
        if colspan || rowspan
          merges << [
            row, col, row + (rowspan || 1)-1, col + (colspan || 1)-1
          ]
        end
      end
    end

    data = { cells: cells, merges: { 0 => merges } }
    from_data(data)
  end

  def self.from_data(data)
    builder = new
    [:merges, :row_heights, :col_widths].each do |d|
      if data[d]
        builder.instance_variable_get("@#{d}".to_sym).merge!(data[d])
      end
    end
    if data[:sheets]
      builder.instance_variable_set(:@sheets, data[:sheets])
    end
    if data[:cells]
      data[:cells].group_by { |c| c[:sheet] }.each do |sheet, cells|
        sheet ||= 0
        if sheet.respond_to?(:to_str)
          builder.set_sheet_by_name(sheet)
        else
          builder.set_sheet(sheet)
        end
        cells.each do |cell|
          # row and col required
          builder.set_cell_value(cell[:row], cell[:col], cell[:value])
          if format = cell[:format]
            builder.add_format_to_cell(cell[:row], cell[:col], format)
          end
        end
      end
    end
    builder
  end

  attr_accessor :name
  attr_reader :book, :sheets
  def initialize
    @cells  = Hash.new { |h,k| h[k] = { value: nil, format: {} } } 
    @sheets      = [] 
    @merges      = Hash.new { |h,k| h[k] = [] }
    @row_heights = Hash.new { |h,k| h[k] = {} }
    @col_widths  = Hash.new { |h,k| h[k] = {} } 
  end

  def _print
    index = current_sheet
    cells = sheet_cells(index)
    0.upto(sheet_rows(cells).last) do |row|
      cols = []
      0.upto(sheet_cols(cells).last) do |col|
        cols << @cells[[index, row, col]][:value]
      end
      puts cols.join("\t")
    end
  end

  def current_sheet
    @current_sheet ||= 0
  end

  def _cells
    @cells
  end

  def to_spreadsheet
    @sheets << 'Sheet 1' if @sheets.empty?

    @book        = Spreadsheet::Workbook.new
    @book_sheets = @sheets.map { |n| book.create_worksheet(name: n) }
    @book_sheets.each_with_index do |sheet, index|
      build_sheet(sheet, index)
    end
    @book
  end 
   
  def add_merge(r1, c1, r2, c2)
    @merges[current_sheet] << [r1, c1 ,r2, c2]
  end

  def set_row_height(row, height)
    @row_heights[current_sheet][row] = Integer(height)
  end

  def set_col_width(col, width)
    @col_widths[current_sheet][col] = Integer(width)
  end

  def add_sheet(name)
    @sheets << name.to_s
  end

  def set_sheet_by_name(name)
    @current_sheet = @sheets.index(name)
  end

  def set_sheet(index)
    @current_sheet = index
  end

  def add_blank_row(row)
    sheet_cols(sheet_cells(current_sheet)).each do |col|
      set_cell_value(row, col, '')
    end
  end

  def set_cell_value(row, col, val)
    @cells[[current_sheet, row, col]][:value] = val
  end

  def add_format_to_cell(row, col, options)
    @cells[[current_sheet, row, col]][:format].merge!(options)
  end

  def add_format_to_row(row, options)
    sheet_cols(sheet_cells(current_sheet)).each do |col|
      add_format_to_cell(row, col, options)
    end
  end

  def add_format_to_col(col, options)
    sheet_rows(sheet_cells(current_sheet)).each do |row|
      add_format_to_cell(row, col, options)
    end
  end

  def add_format_to_box(r1, c1, r2, c2, options)
    (r1..r2).each do |row|
      (c1..c2).each do |col|
        add_format_to_cell(row, col, options)
      end
    end
  end

  #private
  def sheet_cells(index)
    @cells.select { |(sheet,_,_),_| sheet == index }
  end

  def sheet_rows(cells)
    rows = cells.keys.map { |(_,row,_)| row }.sort
    rows.first..rows.last
  end

  def sheet_cols(cells)
    cols = cells.keys.map { |(_,_,col)| col }.sort
    cols.first..cols.last
  end

  def add_each_row(sheet, cells, index)
    0.upto(sheet_rows(cells).last) do |row|
      cols = []
      0.upto(sheet_cols(cells).last) do |col|
        cols << @cells[[index, row, col]][:value]
      end
      sheet.row(row).concat(cols)
    end
  end

  def format_each_row(sheet, cells, index)
    sheet_rows(cells).each do |row|
      sheet_cols(cells).each do |col|
        sheet.row(row).set_format(
          col, 
          Spreadsheet::Format.new(@cells[[index, row, col]][:format])
        )
      end
    end
  end

  def merge_cells(sheet, index)
    @merges[index].each do |points|
      sheet.merge_cells(*points)
    end
  end

  def set_row_heights(sheet, index)
    @row_heights[index].each do |row, height|
      sheet.row(row).height = height
    end
  end

  def set_col_widths(sheet, index)
    @col_widths[index].each do |col, width|
      sheet.column(col).width = width
    end
  end

  def build_sheet(sheet, index)
    cells = sheet_cells(index)

    unless cells.empty?
      add_each_row(sheet, cells, index)
      format_each_row(sheet, cells, index)
      merge_cells(sheet, index)
      set_row_heights(sheet, index)
      set_col_widths(sheet, index)
    end
  end
end
