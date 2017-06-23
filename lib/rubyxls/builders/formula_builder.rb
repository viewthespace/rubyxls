# frozen_string_literal: true
module Rubyxls
  module Builders
    class FormulaBuilder

      attr_reader :cells

      def initialize(**opts)
        @value = opts.fetch(:value, nil)
        @style = opts.fetch(:style, [:default])
        @width = opts.fetch(:width, nil)
        @height = opts.fetch(:height, nil)
        @data_validation = opts.fetch(:data_validation, nil)
        @merge = opts.fetch(:merge, nil)
        @fill = opts.fetch(:fill, :right)
        @num_cells = opts.fetch(:num_cells, nil)
        @row_index = opts.fetch(:row_index, nil)
        @cells = []

        fill_right! if @fill == :right
        fill_down! if @fill == :down
        fill_all! if @fill == :all
      end

      private

      def fill_right!
        raise "Num cells cannot be nil when filling right for formula: #{@value}!" if @num_cells.nil?
        @num_cells.times do
          add_cell!
          shift_value_right!
        end
      end

      def fill_down!
        raise "Row index cannot be nil when filling down for formula: #{@value}!" if @row_index.nil?
        shift_value_down!
        add_cell!
      end

      def fill_all!
        raise "Both num cells & row index cannot be nil when filling all for formula: #{@value}!" if @num_cells.nil? || @row_index.nil?
        shift_value_down!
        fill_right!
      end

      def add_cell!
        @cells << { value: @value.nil? ? nil : @value.clone, style: @style, width: @width, height: @height, data_validation: @data_validation, merge: @merge }
      end

      def shift_value_right!
        tmp_value = @value.dup
        if tmp_value.nil?
         @value = nil
       else
         @value = tmp_value.gsub!(/(?<![\$A-Z])[A-Z]+(?=[\$\d])/) { |column| column.next }
       end
        @value
      end

      def shift_value_down!
        tmp_value = @value.dup
        if tmp_value.nil?
          @value = nil
        else
          @value = tmp_value.gsub!(/(?<=[A-Z])[0-9]+/) { |row| row.to_i + @row_index }
        end
        @value
      end

    end
  end
end
