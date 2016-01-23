class Reporting::Excel2::StyleManagers::DefaultStyleManager

  def initialize(opts={})
    @styles = {}
    @currency_symbol = opts.fetch(:currency_symbol, ActivityLog.new.currency_symbol)
    initialize_default_style!
    initialize_base_styles!
  end

  def retrieve_style_attributes(styles)
    styles = [] if styles.nil?
    styles.each_with_object(@styles[:default].clone) do |style, combined_styles_hash|
      raise("The style :#{style} has not been defined! Please define or create a new type of Style in the styles directory!") if @styles[style].nil?
      combined_styles_hash.deep_merge!(@styles[style])
    end
  end

  private

  def initialize_default_style!
    define_style(style: :default, attributes: { alignment: { horizontal: :left, wrap_text: false, vertical: :bottom }, font_name: 'Arial', sz: 10, border: { style: :none, color: "000000" } })
  end

  def initialize_base_styles!
    define_style(style: :bold, attributes: { b: true })
    define_style(style: :italic, attributes: { i: true })
    define_style(style: :underline, attributes: { u: true })
    define_style(style: :strike, attributes: { strike: true })

    define_style(style: :indent, attributes: { alignment: { indent: 1 } })

    define_style(style: :left_align, attributes: { alignment: { horizontal: :left } })
    define_style(style: :right_align, attributes: { alignment: { horizontal: :right } })
    define_style(style: :center_align, attributes: { alignment: { horizontal: :center } })

    define_style(style: :top_align, attributes: { alignment: { vertical: :top } })
    define_style(style: :bottom_align, attributes: { alignment: { vertical: :bottom } })
    define_style(style: :middle_align, attributes: { alignment: { vertical: :center } })

    define_style(style: :border_right, attributes: { border_right: { style: :thin } })
    define_style(style: :border_left, attributes: { border_left: { style: :thin } })
    define_style(style: :border_top, attributes: { border_top: { style: :thin } })
    define_style(style: :border_bottom, attributes: { border_bottom: { style: :thin } })
    define_style(style: :border_all, attributes: { border_right: { style: :thin }, border_left: { style: :thin }, border_top: { style: :thin }, border_bottom: { style: :thin } })

    define_style(style: :number, attributes: { format_code: '#,##0_ ;(#,##0)' })
    define_style(style: :decimal, attributes: { format_code: '#,##0.00_ ;(#,##0.00)' })
    define_style(style: :date, attributes: { format_code: 'MM/D/YY' })
    define_style(style: :time, attributes: { format_code: "[#{@currency_symbol}-409]h:mm AM/PM;@" })
    define_style(style: :currency, attributes: { format_code: "#{@currency_symbol}#,##0;(#{@currency_symbol}#,##0)" })
    define_style(style: :currency_precision, attributes: { format_code: "#{@currency_symbol}#,##0.00;(#{@currency_symbol}#,##0.00)" })
    define_style(style: :percent, attributes: { format_code: '0.00%;(0.00%)' })

    define_style(style: :wrap_text, attributes: { alignment: { wrap_text: true } })
  end

  def define_style(style:, attributes: {})
    @styles[style] = attributes
  end

end
