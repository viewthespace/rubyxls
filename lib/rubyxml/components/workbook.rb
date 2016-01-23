class Reporting::Excel2::Workbooks::DefaultWorkbook < SimpleDelegator

  attr_reader :sheets

  def initialize(**opts)
    @package = Axlsx::Package.new
    @package.use_shared_strings = true
    @name = opts.fetch(:name, "default_workbook")
    @style_codes = {}
    @sheets = []

    super(@package.workbook)
    build_sheets!
    add_sheets!
  end

  def to_stream
    @package.to_stream
  end

  def filename
    "#{@name.parameterize}.xlsx"
  end

  private

  def build_sheets!
    @sheets << Reporting::Excel2::Sheets::DefaultSheet.new
  end

  def add_sheets!
    @sheets.each do |sheet|
      @data_validations = {}
      add_sheet(sheet.build_options)
      add_rows(sheet.build_rows)
      add_data_validations
      add_charts(sheet.build_charts)
      add_width_normalization(sheet.build_width_normalization)
    end
  end

  def add_sheet(options)
    add_worksheet(options) do |sheet|
      sheet.sheet_view.show_grid_lines = false
      sheet.page_setup.fit_to(width: 1, height: 999)
      @sheet = sheet
    end
  end

  def add_rows(rows)
    rows.each do |row|
      values = []
      widths = []
      styles = []
      height = nil
      row.each do |cell|
        values << cell[:value]
        widths << cell[:width]
        styles << retrieve_style_code(cell[:style])
        height = cell[:height] if cell[:height]
        merge_cell(cell) if cell[:merge]
        @data_validations["#{cell[:column]}#{cell[:row]}"] = cell[:data_validation] if cell[:data_validation]
      end
      @sheet.add_row(values, { style: styles, widths: widths, height: height })
    end
  end

  def retrieve_style_code(style_attributes)
    hash_attributes = style_attributes.clone
    @style_codes[hash_attributes] ||= styles.add_style(style_attributes)
  end

  def merge_cell(cell)
    case cell[:merge]
    when :right
      @sheet.merge_cells("#{cell[:column]}#{cell[:row]}:#{cell[:column].next}#{cell[:row]}")
    when :down
      @sheet.merge_cells("#{cell[:column]}#{cell[:row]}:#{cell[:column]}#{cell[:row] + 1}")
    end
  end

  def add_data_validations
    @data_validations.each do |data_validation_cell, data_validation_values|
      start_column = "A"
      end_column = start_column.clone
      (data_validation_values.size - 1).times { end_column.next! }
      @sheet.add_row(data_validation_values, { style: Array.new(data_validation_values.size, retrieve_style_code({ fg_color: "FFFFFF", bg_color: "FFFFFF" })) })
      @sheet.add_data_validation(data_validation_cell, { type: :list, formula1: "$#{start_column}$#{@sheet.rows.count}:$#{end_column}$#{@sheet.rows.count}", showDropdown: false })
    end
  end

  def add_charts(charts)
    charts.each do |chart|
      @sheet.add_chart(chart.type, start_at: chart.start_at, end_at: chart.end_at, rotX: chart.rot_x, rotY: chart.rot_y) do |added_chart|
        added_chart.title = chart.title
        added_chart.bar_dir = chart.bar_dir if added_chart.respond_to?(:bar_dir)
        added_chart.show_legend = chart.show_legend
        added_chart.catAxis.title = chart.cat_axis_title if added_chart.respond_to?(:catAxis)
        added_chart.valAxis.title = chart.val_axis_title if added_chart.respond_to?(:valAxis)
        added_chart.catAxis.gridlines = chart.cat_axis_gridlines if added_chart.respond_to?(:catAxis)
        added_chart.valAxis.gridlines = chart.val_axis_gridlines if added_chart.respond_to?(:valAxis)
        chart.series.each do |series|
          added_chart.add_series(data: @sheet[series[:data]], labels: @sheet[series[:labels]], colors: series[:colors])
        end
      end
    end
  end

  def add_width_normalization(custom_widths: [], default_width: nil)
    if @sheet.rows.size > 0
      row_values = Array.new(@sheet.cols.count, nil)
      row_widths = Array.new(@sheet.cols.count, default_width)
      custom_widths.reverse.each do |custom_width|
        row_widths.unshift(custom_width)
        row_widths.pop
      end
      @sheet.add_row(row_values, { widths: row_widths })
    end
  end

end
