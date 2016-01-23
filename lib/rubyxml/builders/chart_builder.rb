class Reporting::Excel2::Builders::ChartBuilder

  attr_reader :type
  attr_reader :start_at
  attr_reader :end_at
  attr_reader :title
  attr_reader :rot_x
  attr_reader :rot_y
  attr_reader :show_legend
  attr_reader :bar_dir
  attr_reader :cat_axis_title
  attr_reader :val_axis_title
  attr_reader :cat_axis_gridlines
  attr_reader :val_axis_gridlines
  attr_reader :series

  def initialize(**opts)
    @type = opts.fetch(:type, Axlsx::Bar3DChart)
    @start_at = opts.fetch(:start_at, "A1")
    @end_at = opts.fetch(:end_at, "A1")
    @title = opts.fetch(:title, " ")
    @rot_x = opts.fetch(:rot_x, 30)
    @rot_y = opts.fetch(:rot_y, 20)
    @show_legend = opts.fetch(:show_legend, false)
    @bar_dir = opts.fetch(:bar_dir, :col)
    @cat_axis_title = opts.fetch(:cat_axis_title, " ")
    @val_axis_title = opts.fetch(:val_axis_title, " ")
    @cat_axis_gridlines = opts.fetch(:cat_axis_gridlines, false)
    @val_axis_gridlines = opts.fetch(:val_axis_gridlines, false)
    @series = opts.fetch(:series, [])
  end

end
