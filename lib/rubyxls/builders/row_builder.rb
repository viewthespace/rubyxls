module Rubyxls
  module Builders
    class RowBuilder

      attr_reader :rows

      def initialize(cells)
        @cells = cells
        @bottom_row_bound = 0
        @right_column_bound = 0
        expand_bottom_row_bound
        expand_right_column_bound
        @rows = Array.new(@bottom_row_bound) { Array.new(@right_column_bound) { { value: nil, style: [] } } }
        build_rows!
      end

      private

      def expand_bottom_row_bound
        @cells.each do |cell|
          @bottom_row_bound = cell[:row] if cell[:row] > @bottom_row_bound
        end
      end

      def expand_right_column_bound
        @cells.each do |cell|
          @right_column_bound = convert_cell_column_to_integer(cell[:column]) if convert_cell_column_to_integer(cell[:column]) > @right_column_bound
        end
      end

      def convert_cell_column_to_integer(cell_column)
        offset = "A".ord - 1
        cell_column.chars.inject(0) { |sum,char| sum * 26 + char.ord - offset }
      end

      def build_rows!
        @cells.each { |cell| @rows[retrieve_cell_row_index(cell[:row])][retrieve_cell_column_index(cell[:column])] = cell }
      end

      def retrieve_cell_row_index(cell_row)
        cell_row - 1
      end

      def retrieve_cell_column_index(cell_column)
        convert_cell_column_to_integer(cell_column) - 1
      end

    end
  end
end
