module Rubyxls
  class Sheet

    PAGE_SETUP = { fit_to_width: 1, orientation: :landscape }
    PRINT_OPTIONS = { grid_lines: false }
    PAGE_MARGINS = { top: 0.3, left: 0.3, bottom: 0.3, right: 0.3, header: 0, footer: 0 }

    def initialize(**opts)
      @sheet_name = opts.fetch(:sheet_name, "Default Sheet")
      @style_manager = opts.fetch(:style_manager, Rubyxls::StyleManager.new)
    end

    def build_options
      { page_setup: PAGE_SETUP, print_options: PRINT_OPTIONS, page_margins: PAGE_MARGINS, name: @sheet_name }
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

    def build_cells
      Rubyxls::Builders::CellBuilder.new(model_data_rows: Rubyxls::ViewModel.new(title_row: true, header_row: true, additional_rows: 1, total_row: true).data_rows, start_row: 2, start_column: "B").cells
    end

  end
end
