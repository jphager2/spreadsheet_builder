module SpreadsheetBuilder
  class CssParser
    attr_reader :cache, :rules

    def initialize
      reset
    end

    def reset(level = :parser)
      method = "reset_#{level}"

      __send__(method)
    end

    def format_from_node(node)
      format_from_klass_tree(klass_tree_from_node(node))
    end

    private
    def accepted_keys
      # TODO Keep these in a config
      keys = %w{ color background-color font-size font-weight text-align border border-width border-style border-color height width }
      dirs = %w{ top bottom left right }
      types = %w{ width style color }
      dirs.each do |dir|
        keys << "border-#{dir}"
        types.each { |type| keys << "border-#{dir}-#{type}" }
      end
      keys
    end

    def klass_tree_from_node(base)
      tree = (base.ancestors.reverse.drop(1) << base)
      tree.map { |n, t| klass_node_from_node(n) }
    end

    def format_from_klass_tree(klass)
      # klass is uniq to each node (because of first-child, nth-child, etc)
      # so caching with the class is useless
      # TODO find a better way to cache that works
      if @cache[klass]
        format = @cache[klass]
      else
        declarations = declarations_from_klass_tree(klass)
        format = format_from_declarations(declarations)
        @cache[klass] = format
      end

      format || {}
    end

    def klass_node_from_node(node)
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

    def find_rules(tree)
      @rules.each_with_object(Hash.new {|h,k| h[k] = []}) do |rule,rules_found|
        found = rule.find_selector(tree)
        rules_found[found.length] << rule if found
      end
    end

    def reset_rules(parser = @parser)
      @rules = []
      parser.each_rule_set(:all) { |r| 
        @rules << SpreadsheetBuilder::CssRule.new(r) 
      }
      @rules.delete_if { |rule| 
        rule.each_declaration do |key,_|
          rule.remove_declaration!(key) unless accepted_keys.include?(key)
        end
        rule.declarations.empty?
      }
    end

    def translate_declaration(key, val)
      f = TRANSLATIONS[key].call(val) if key && val
      f || {}
    end

    def format_from_declarations(declarations)
      declarations.delete_if { |_,v| v.nil? || v.empty? }

      declarations.each_with_object({}) { |(k,v), format|
        format.merge!(translate_declaration(k,v[:value]))
      }
    end

    def declarations_from_klass_tree(klass)
      rules = find_rules(klass).sort_by { |specificity,_| specificity }
      declarations = rules.map { |_,rules| 
        rules.inject({}) { |dec, r| dec.merge(r.declarations) }
      }.inject({},&:merge)
    end

    def reset_cache
      @cache = {}
    end

    def reset_parser
      parser = CssParser::Parser.new
      # TODO load these files from a config
      parser.load_uri!("file://#{Dir.pwd}/test2.css")
      # TODO or even better parse the html doc for spreadsheet links 
      # and load those
      #parser.load_uri!("file://#{Dir.pwd}/test2.css")

      # Explicity reset rules to avoid infinite loop or bad data
      reset_rules(parser)
      reset_cache
      @parser = parser
    end
  end
end
