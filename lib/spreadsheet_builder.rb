require 'nokogiri'
require 'spreadsheet'

class SpreadsheetBuilder

  H1 = { weight: :bold }
  SOMETHING_ELSE = { color: :red }
  TEXT_CENTER = { horizontal_align: :center }

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

  def self._format_from_klass(klass)
    merge(
      *klass.map { |k| 
        begin
          "#{to_s}::#{k.to_s.underscore.upcase}".constantize
        rescue NameError
          {}
        end
      }
    )
  end

  def self._add_td_to_cells(row, col, td, tr_klass, cells)
    found = cells.find { |cell| cell[:row] == row && cell[:col] == col}
    unless found 
      klass  = td.attributes["class"] && td.attributes["class"].value
      klass  = klass.to_s.split(/ /) + tr_klass
      format = _format_from_klass(klass)
      cells << { 
        row: row, 
        col: col, 
        value: td.text.strip, 
        format: format, 
        path: td.css_path 
      } 
    else
      _add_td_to_cells(row, col + 1, td, tr_klass, cells)
    end
  end

  def self.from_html(html)
    # need to check for attributes colspan and row span
    cells       = [] 
    merges      = []
    col_widths  = {}
    row_heights = {}

    doc      = Nokogiri::HTML(html)
    tb       = doc.css('table').first 
    tb_klass = tb.attributes["class"] && tb.attributes["class"].value
    tb_klass = tb_klass.to_s.split(/ /)

    doc.css('tr').each_with_index do |tr, row|
      tr_klass = tr.attributes["class"] && tr.attributes["class"].value
      tr_klass = tr_klass.to_s.split(/ /) + tb_klass

      tr.css('td, th').each_with_index do |td, col|
         
        rowheight = td.attributes["rowheight"]
        colwidth  = td.attributes["colwidth"]
        rowspan   = td.attributes["rowspan"]
        colspan   = td.attributes["colspan"]

        rowheight &&= rowheight.value.to_i
        colwidth  &&= colwidth.value.to_i
        rowspan   &&= rowspan.value.to_i
        colspan   &&= colspan.value.to_i

        _add_td_to_cells(row, col, td, tr_klass, cells)
        if colspan
          (1..colspan-1).each {|t| 
            _add_td_to_cells(row, col+t, td, tr_klass, cells)
          }
        end
        if rowspan
          (1..rowspan-1).each {|t| 
            _add_td_to_cells(row+t, col, td, tr_klass, cells)
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
