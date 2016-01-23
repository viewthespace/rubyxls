require 'spec_helper'

describe Reporting::Excel2::Builders::ChartBuilder do
  let(:chart_builder) { Reporting::Excel2::Builders::ChartBuilder.new }

  describe '#type' do

    it 'initializes with the correct default value' do
      expect(chart_builder.type).to eq(Axlsx::Bar3DChart)
    end

  end

  describe '#start_at' do

    it 'initializes with the correct default value' do
      expect(chart_builder.start_at).to eq("A1")
    end

  end

  describe '#end_at' do

    it 'initializes with the correct default value' do
      expect(chart_builder.end_at).to eq("A1")
    end

  end

  describe '#title' do

    it 'initializes with the correct default value' do
      expect(chart_builder.title).to eq(" ")
    end

  end

  describe '#rot_x' do

    it 'initializes with the correct default value' do
      expect(chart_builder.rot_x).to eq(30)
    end

  end

  describe '#rot_y' do

    it 'initializes with the correct default value' do
      expect(chart_builder.rot_y).to eq(20)
    end

  end

  describe '#show_legend' do

    it 'initializes with the correct default value' do
      expect(chart_builder.show_legend).to be(false)
    end

  end

  describe '#bar_dir' do

    it 'initializes with the correct default value' do
      expect(chart_builder.bar_dir).to eq(:col)
    end

  end

  describe '#cat_axis_title' do

    it 'initializes with the correct default value' do
      expect(chart_builder.cat_axis_title).to eq(" ")
    end

  end

  describe '#val_axis_title' do

    it 'initializes with the correct default value' do
      expect(chart_builder.val_axis_title).to eq(" ")
    end

  end

  describe '#cat_axis_gridlines' do

    it 'initializes with the correct default value' do
      expect(chart_builder.cat_axis_gridlines).to be(false)
    end

  end

  describe '#val_axis_gridlines' do

    it 'initializes with the correct default value' do
      expect(chart_builder.val_axis_gridlines).to be(false)
    end

  end

  describe '#series' do

    it 'initializes with the correct default value' do
      expect(chart_builder.series).to eq([])
    end

  end

end