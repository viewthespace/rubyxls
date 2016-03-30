module Rubyxls

  def self.generate_default_report
    Report.new
  end

  def self.write_report_to_tmp(report)
    file_path = report_tmp_file_path(report)
    FileUtils.mkdir_p(tmp_dir_path) unless File.directory?(tmp_dir_path)
    File.delete(file_path) if File.exist?(file_path)
    File.open(file_path, 'w') do |file|
      file << report.download!.read
    end
  end

  def self.open(report)
    `open "#{report_tmp_file_path(report)}"`
  end

  private

  def self.tmp_dir_path
    File.expand_path('../../tmp', __FILE__)
  end

  def self.report_tmp_file_path(report)
    tmp_dir_path + "/#{report.file_name}.#{report.file_extension}"
  end

end

require 'rubyxls/version'

require 'rubyxls/report'
require 'rubyxls/sheet'
require 'rubyxls/style_manager'
require 'rubyxls/view_model'
require 'rubyxls/workbook'

require 'rubyxls/builders/cell_builder'
require 'rubyxls/builders/chart_builder'
require 'rubyxls/builders/formula_builder'
require 'rubyxls/builders/row_builder'
