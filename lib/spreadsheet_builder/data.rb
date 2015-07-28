module SpreadsheetBuilder
  def self.from_data(data)
    builder = Builder.new
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
end
