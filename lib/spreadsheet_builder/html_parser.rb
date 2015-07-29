module SpreadsheetBuilder
  class HtmlParser
    def self.from_slim(file, options = {}, context = self, &block)
      html = Slim::Template.new(file, options).render(context, &block)
      new(html)
    end

    def self.from_erb(file)
      html     = File.read(file)
      template = ERB.new(html)
      html     = template.result

      new(html)
    end

    def initialize(html)
      @html = html
      @css  = SpreadsheetBuilder::CssParser.new
    end

    def build(force_level = :none)
      SpreadsheetBuilder.from_data(to_data)
    end

    # TODO clean this up
    def to_data(force_level = :none)
      reset_css_parser if force_level == :parser
      reset_css_cache  if force_level == :cache
      reset_css_rules  if force_level == :rules

      # need to check for attributes colspan and row span
      cells       = [] 
      merges      = []
      col_widths  = {}
      row_heights = {}

      doc = Nokogiri::HTML(@html)
      tb  = doc.css('table').first 

      # ignoring specified formats for anything other than table tr td/th
      tb_format = @css.format_from_node(tb) 

      doc.css('tr').each_with_index do |tr, row|
        tr_format = tb_format.merge(@css.format_from_node(tr))

        tr.css('td, th').each_with_index do |td, col|
           
          rowheight = td.attributes["rowheight"]
          colwidth  = td.attributes["colwidth"]
          rowspan   = td.attributes["rowspan"]
          colspan   = td.attributes["colspan"]

          rowheight &&= rowheight.value.to_i
          colwidth  &&= colwidth.value.to_i
          rowspan   &&= rowspan.value.to_i
          colspan   &&= colspan.value.to_i

          add_td_to_cells(row, col, td, tr_format, cells)
          if colspan
            (1..colspan-1).each {|t| 
              add_td_to_cells(row, col+t, td, tr_format, cells)
            }
          end
          if rowspan
            (1..rowspan-1).each {|t| 
              add_td_to_cells(row+t, col, td, tr_format, cells)
            }
          end
          if colspan || rowspan
            merges << [
              row, col, row + (rowspan || 1)-1, col + (colspan || 1)-1
            ]
          end
        end
      end

      { cells: cells, merges: { 0 => merges } }
    end

    private
    # TODO Document
    def add_td_to_cells(row, col, td, tr_format, cells)
      found = cells.find { |cell| cell[:row] == row && cell[:col] == col}
      unless found 
        td_format = tr_format.merge(@css.format_from_node(td))
        cells << { 
          row: row, 
          col: col, 
          value: td.text.strip, 
          format: td_format, 
          path: td.css_path 
        } 
      else
        add_td_to_cells(row, col + 1, td, tr_format, cells)
      end
    end
  end
end
