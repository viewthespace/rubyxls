class Rubyxls::Builders::CellBuilder

  attr_reader :cells

  def initialize(**opts)
    @model_data_rows = opts.fetch(:model_data_rows)
    @start_row = opts.fetch(:start_row, 1)
    @start_column = opts.fetch(:start_column, "A")
    @cells = []
    build_cells!
  end

  private

  def build_cells!
    assign_row_column!
    @cells = @model_data_rows.flatten
  end

  def assign_row_column!
    @model_data_rows.each_with_index do |data_row, table_row_index|
      data_row.each_with_index do |data_cell, table_column_index|
        data_cell[:row] = @start_row + table_row_index
        data_cell[:column] = retrieve_cell_column_letter(@start_column, table_column_index)
      end
    end
  end

  def retrieve_cell_column_letter(start_column, table_column_index)
     cell_column = start_column.clone
     table_column_index.times { cell_column.next! }
     cell_column
  end

end
