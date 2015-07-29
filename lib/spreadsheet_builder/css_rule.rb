module SpreadsheetBuilder
  class CssRule
    def initialize(rule)
      @rule = rule
    end

    def declarations
      @rule.instance_variable_get(:@declarations)
    end

    def selectors 
      @rule.selectors.map { |s| 
        s = s.split(/[\s>]/).map { |node| node.split('.') }
        s.each do |node| 
          n = node.length
          node[1...n] = node[1...n].map { |k| '.' + k }
          node.delete_if { |k| k.empty? || k == "." }
        end 
      }.sort_by(&:length)
    end

    def find_selector(tree) 
      selectors.find { |s| 
        next unless tree.last & s.last == s.last

        s[0..-2].reverse.inject(tree.length - 2) do |i, s_node|
          break false unless i
          tree[0..i].index { |kt_node| kt_node & s_node == s_node }
        end
      }
    end

    private
    def method_missing(method, *attrs, &block)
      if @rule.respond_to?(method)
        @rule.__send__(method, *attrs, &block)
      else
        super
      end
    end
  end
end
