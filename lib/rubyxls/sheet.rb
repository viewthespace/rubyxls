module Rubyxls
  class Sheet

    FIT_TO_WIDTH = 1
    ORIENTATION = :landscape
    GRID_LINES = false
    PAGE_MARGIN_TOP = 0.3
    PAGE_MARGIN_LEFT = 0.3
    PAGE_MARGIN_BOTTOM = 0.3
    PAGE_MARGIN_RIGHT = 0.3
    HEADER = 0.0
    FOOTER = 0.0
    LETTER_PAPER_SIZE = 1

    attr_reader :sheet_name

    def initialize(**opts)
      @sheet_name = opts.fetch(:sheet_name, "Default Sheet")
      @style_manager = opts.fetch(:style_manager, Rubyxls::StyleManager.new)
      @paper_size = opts.fetch(:paper_size, LETTER_PAPER_SIZE)
    end

    def build_options(*taken_names)
      { page_setup: build_page_setup,
        print_options: build_print_options,
        page_margins: build_page_margins,
        name: unique_sheet_name(sanitize_sheet_name(@sheet_name), taken_names) }
    end

    def build_rows
      Rubyxls::Builders::RowBuilder.new(build_cells).rows.each do |row|
        row.each do |cell|
          cell[:style] = @style_manager.retrieve_style_attributes(cell[:style])
        end
      end
    end

    def build_charts
      []
    end

    def build_width_normalization
      {}
    end

    private

    def build_page_setup
      { fit_to_width: FIT_TO_WIDTH,
        orientation: ORIENTATION,
        paper_size: @paper_size }
    end

    def build_print_options
      { grid_lines: GRID_LINES }
    end

    def build_page_margins
      { top: PAGE_MARGIN_TOP,
        left: PAGE_MARGIN_LEFT,
        bottom: PAGE_MARGIN_BOTTOM,
        right: PAGE_MARGIN_RIGHT,
        header: HEADER,
        footer: FOOTER }
    end

    def build_cells
      Rubyxls::Builders::CellBuilder.new(model_data_rows: Rubyxls::ViewModel.new(title_row: true, header_row: true, additional_rows: 1, total_row: true).data_rows, start_row: 2, start_column: "B").cells
    end

    def sanitize_sheet_name(sheet_name)
      sheet_name.gsub(/[:\[\]\/\\\*\?]/, '-')
    end

    def unique_sheet_name(sheet_name, taken_names, index=1)
      sliced_name = sheet_name.byteslice(0, 31)
      return sliced_name unless taken_names.include?(sliced_name)

      sheet_name = "#{sheet_name.byteslice(0, 28)}(#{index})"
      unique_sheet_name(sheet_name, taken_names, index + 1)
    end

  end
end
