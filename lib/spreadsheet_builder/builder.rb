module SpreadsheetBuilder
  class Builder

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

      @book = Spreadsheet::Workbook.new
      
      CUSTOM_PALETTE.each do |name, color|
        @book.set_custom_color(name, color.r, color.g, color.b)
      end

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
end
