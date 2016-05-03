# Rubyxls

Generate XLS files using ruby. Rubyxls provides a simple DSL :o express anything from simple ruby strings to complex Active Record models.
Support for multi-sheet workbooks, chart generation, and formula based cells included.

**This would not have been possible without the hard work of the Axlsx team. Thanks for creating the foundation
that Rubyxls is built on top of! Please be sure to check out their repo here: https://github.com/randym/axlsx.**

## Installation

Add this line to your application's Gemfile:

```ruby
gem 'rubyxls'
```

And then execute:

    $ bundle

Or install it yourself as:

    $ gem install rubyxls

## Usage

### Components
After installing, Rubyxls is easy to use.

Rubyxls has the same components of an actual excel spreadsheet, abstracting away the messy details.

I like to think of the components in the following way:
- **Report**: has_many workbooks (yes, a report can have many workbooks resulting in a zip file of all the workbooks)
- **Workbook**: has_many sheets (when creating a new workbook, you are essentially creating a new xls file)
- **Sheet**: has_many view models (tables)
- **View Model (Table)**: this is where all of the data goes. More on ViewModels to come...
- **Style Manager**: create re-usable styles to make sure your reports are beautiful and consistent.

**Please don't be fooled by the has_many term listed above. This has nothing to do with ActiveRecord or any database relation. As a ruby/rails developer, I have just gotten in the habit of using this association terminology to describe any kind of relationship : )**

In order to create your very own excel report, all you have to do is inherit each component's functionality that Rubyxls gives you.

**Note: Rubyxls is built with the intention that you will include each of the listed components in your excel project, allowing each component to play nicely with each other.**

#### Report
```ruby
class YourVeryOwnReport < Rubyxls::Report

  def initialize(whatever_data_you_need)
    @data = whatever_data_you_need
    super()                               # Call #super with no parameters
  end

  private

  def build_workbooks!
    @workbooks << YourVeryOwnWorkbook.new(whatever_data_you_need)
  end
end
```
- **Inherit from Rubyxls::Report** - this will give you all the functionality a report needs to build itself.
- **Call super() in #initialize** - this will create an instance variable `@workbooks` to hold all of the beautfiul workbooks you are about to create!
- **#build_workbooks!** - append any workbooks that you create into `@workbooks` so that your report "has_many" workbooks just like the real excel!

#### Workbook
```ruby
class YourVeryOwnWorkbook < Rubyxls::Workbook

  def initialize(whatever_data_you_neeed)
    @data = whatever_data_you_need
    super(name: "your_very_own_workbook")   # Call #super passing in only a name parameter
  end

  private

  def build_sheets!
    @sheets << YourVeryOwnSheet.new(whatever_data_you_need)
  end

end
```
- **Inherit from Rubyxls::Workbook** - get all of that groovy (not the programming language) functionality that Rubyxls gives your Workbook.
- **Call super() in #initialize passing in a "name" parameter** - this will create an instance variable `@sheets` to hold all of your sheets and will name the workbook you are creating.
- **#build_sheets!** - append any sheets that you create into `@sheets` so that your workbook "has_many" sheets!

#### Sheet
```ruby
class YourVeryOwnSheet < Rubyxls::Sheet

  def initialize(whatever_data_you_need)
    @data = whatever_data_you_need
    super(sheet_name: "YourVeryOwnSheet")    # Call #super passing in only a sheet_name parameter
  end

  private

  def build_cells
    Rubyxls::Builders::CellBuilder.new(model_data_rows: YourVeryOwnViewModel.new(title: "Your Very Own View Model").data_rows, start_row: 1, start_column: "A").cells
    .+ Rubyxls::Builders::CellBuilder.new(model_data_rows: AnotherViewModel.new(title: "Another View Model").data_rows, start_row: 4, start_column: "L").cells
  end

end
```
- **Inherit from Rubyxls::Sheet** - gotta get that sheet functionality
- **Call super() in #initialize passing in a "name" parameter** - this will give your sheet a name
- **#build_cells!** - this is the dumping ground for all of your data. The CellBuilder, which takes in a view model as a parameter will break apart your view models / tables into invidual
cells and organize them properly. Don't forget to specify where on the sheet you would like your view models to be drawn using `start_row` and `start_column`!
Behind the scenes, we are taking all of your modularized view models, breaking them apart into individual cells, throwing all of those cells into one LARGE pot and re-arranging them in the
proper order. Cool, huh!

#### ViewModel
```ruby
class YourVeryOwnViewModel < Rubyxls::ViewModel

  def initialize(whatever_data_you_need)
    @data = whatever_data_you_need
    super()                             # Call #super passing in no parameters
  end

  private

  def build_data_rows!
    @data_rows << [
                   { value: "Title Row", style: [:bold] }
                  ]
  end

end
```
- **Inherit from Rubyxls::ViewModel**
- **Call super() in #initialize passing**
- **#build_data_rows!** - this is where you decide what data goes into the workbook and how it looks. Let's break it down by describing the ruby version of each excel component that makes up a row.
  ```
  | Ruby  | Excel | Desciption                                             |
  | -----  -------  ------------------------------------------------------ |
  | Array | Row   | Each array has lots of individual cells or ruby hashes |
  | Hash  | Cell  | Each cell has a value and a style                      |
  ```

  `@data_rows` starts as an empty array that is defined when you call `super()` in the initialize method. The whole idea behind the View Model is to populate `@data_rows` with ruby arrays, each
  representing a row in excel. Each ruby array is filled with ruby hashes, representing an individual cell.

  Each individual cell is structured as follows:
  ```ruby
  { value: <any value>, style: [<list of styles to apply to this cell>] }
  ```
  A list of default styles is provided below!

### Styling
#### Default Styles
- :bold
- :italic
- :underline
- :strike
- :indent
- :left_align
- :right_align
- :center_align
- :top_align
- :bottom_align
- :middle_align
- :border_right
- :boder_left
- :boder_top
- :boder_bottom
- :boder_all
- :number
- :decimal
- :date
- :time
- :currency
- :currency_precision
- :percent
- :wrap_text

#### Custom Styles
``` ruby
class YourVeryOwnCustomStyleManager < Rubyxls::StyleManager

  GREY = 'CCCCCC'

  def initialize
    super                       # Call super
  end

  private

  def initialize_base_styles!
    define_style(style: :medium, attributes: { sz: 12 })
    define_style(style: :large, attributes: { sz: 14 })
    define_style(style: :total_row, attributes: { b: true, bg_color: GREY })
    super
  end

end
```
- **Inherit from Rubyxls::StyleManager**
- **Call super in #initialize passing** - in order to get all of those default styles that Rubyxls provides.
- **#initialize_base_styles!** - define your own styles according the the excel specification. Each style only needs a set of attributes.

Now you can apply your newly defined styles inside of any sheet.

**Note: in order to use user-defined styles, you must specify which style manager you would like to use when creating your sheet.**
```ruby
class YourVeryOwnSheet

  def initialize(data)
    @data = data
    super(sheet_name: "YourVeryOwnSheet", style_manager: YourVeryOwnCustomStyleManager.new)
  end

end
```
Now you will have access to your user-defined styles and the default styles that ship with Rubyxls!

## Development

After checking out the repo, run `bin/setup` to install dependencies. Then, run `rake spec` to run the tests. You can also run `bin/console` for an interactive prompt that will allow you to experiment.

To install this gem onto your local machine, run `bundle exec rake install`. To release a new version, update the version number in `version.rb`, and then run `bundle exec rake release`, which will create a git tag for the version, push git commits and tags, and push the `.gem` file to [rubygems.org](https://rubygems.org).

## Contributing

Bug reports and pull requests are welcome on GitHub at https://github.com/viewthespace/rubyxls. This project is intended to be a safe, welcoming space for collaboration, and contributors are expected to adhere to the [Contributor Covenant](http://contributor-covenant.org) code of conduct.


## License

The gem is available as open source under the terms of the [MIT License](http://opensource.org/licenses/MIT).

