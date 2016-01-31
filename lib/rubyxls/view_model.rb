module Rubyxls
  class ViewModel

    attr_reader :data_rows

    def initialize(**opts)
      @data_rows = []
      @data_rows_count = opts.fetch(:rows_count, nil)

      build_title_row!
      build_header_row!
      build_data_rows!
      build_additional_rows! unless @data_rows_count.nil?
      build_total_row!
    end

    private

    def build_title_row!
      []
    end

    def build_header_row!
      []
    end

    def build_data_rows!
      []
    end

    def build_additional_row!
      []
    end

    def build_total_row!
      []
    end

    def build_additional_rows!
      until @data_rows.size == @data_rows_count do
        build_additional_row!
      end
    end

    def add_empty_cell(*style)
      { value: nil, style: style }
    end

    def limit_data_to_data_rows_count(data)
      @data_rows_count.nil? ? data : data[0...calculate_data_rows_remaining]
    end

    def calculate_data_rows_remaining
      @data_rows_count - @data_rows.size
    end
    
  end
end
