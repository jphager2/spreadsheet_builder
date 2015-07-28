module Spreadsheet
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

    def new(html)
      @html = html
    end

    # TODO clean this up
    def builder(force_level = :none)
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
      tb_klass  = klass_tree_for_node(tb)
      tb_format = format_from_klass_tree(tb_klass) 

      doc.css('tr').each_with_index do |tr, row|
        tr_klass  = klass_tree_for_node(tr)
        tr_format = tb_format.merge(format_from_klass_tree(tr_klass))

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

      data = { cells: cells, merges: { 0 => merges } }
      from_data(data)
    end

    private
    def translate_declaration(key, val)
      f = TRANSLATIONS[key].call(val) if key && val
      f || {}
    end

    def format_for_declarations(declarations)
      declarations.delete_if { |_,v| v.nil? || v.empty? }

      declarations.each_with_object({}) { |(k,v), format|
        format.merge!(translate_declaration(k,v[:value]))
      }
    end

    def format_from_klass_tree(klass)
      if _css_cache[klass]
        format = css_cache[klass]
      else
        declarations = _declarations_from_klass_tree(klass)
        format = _format_for_declarations(declarations)
        css_cache[klass] = format
      end

      format || {}
    end

    def _declarations_from_klass_tree(klass)
      rules = _find_rules(klass).sort_by { |specificity,_| specificity }
      declarations = rules.map { |_,rules| 
        rules.inject({}) { |dec, r| 
          dec.merge(r.instance_variable_get(:@declarations))
        }
      }.inject({},&:merge)
    end

    def _css_rules
      @css_rules || reset_css_rules
    end

    def _css_cache
      @css_cache || reset_css_cache
    end

    def _css_parser
      @css_parser || reset_css_parser
    end

    def _reset_css_rules(parser = _css_parser)
      @css_rules = []
      parser.each_rule_set(:all) { |r| @css_rules << r }
      @css_rules
    end

    def _reset_css_cache
      @css_cache = {}
    end

    def _reset_css_parser
      parser = CssParser::Parser.new
      # TODO load these files from a config
      parser.load_uri!("file://#{Dir.pwd}/test.css")
      # TODO or even better parse the html doc for spreadsheet links 
      # and load those
      #parser.load_uri!("file://#{Dir.pwd}/test2.css")

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
      reset_css_rules(parser)
      css_rules.each do |rule|
        rule.each_declaration do |key,_|
          rule.remove_declaration!(key) unless accepted_keys.include?(key)
        end
      end
       css_rules.delete_if { |rule| 
         rule.instance_variable_get(:@declarations).empty?
       }

      reset_css_cache
      @css_parser = parser
    end

    def add_td_to_cells(row, col, td, tr_format, cells)
      found = cells.find { |cell| cell[:row] == row && cell[:col] == col}
      unless found 
        td_klass  = klass_tree_for_node(td)
        td_format = tr_format.merge(format_from_klass_tree(td_klass))
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

    # TODO support :first-child, :nth-child(n|even|odd), :last-child, :first-of-kind, :last-of-kind, ... ?
    def klass_node_for_node(node)
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

    def klass_tree_for_node(base)
      tree = (base.ancestors.reverse.drop(1) << base)
      tree.map { |n, t| klass_node_for_node(n) }
    end

    def find_rules(klass_tree)
      rules = Hash.new { |h,k| h[k] = [] } 
      css_rules.each do |rule|
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
  end
end
