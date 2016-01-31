require 'spec_helper'

describe Rubyxls::Components::Sheet do
  let(:default_sheet) { Rubyxls::Components::Sheet.new }

  describe '#build_options' do

    subject { default_sheet.build_options }

    it 'builds the correct options' do
      expect(subject).to eq({
                              page_setup: { fit_to_width: 1, orientation: :landscape },
                              print_options: { grid_lines: false },
                              page_margins: { top: 0.3, left: 0.3, bottom: 0.3, right: 0.3, header: 0, footer: 0 },
                              name: " Sheet"
                            })
    end

  end

  describe '#build_rows' do

    subject { default_sheet.build_rows }

    it 'builds the correct number of rows' do
      expect(subject.size).to eq(0)
    end

  end

  describe '#build_charts' do

    subject { default_sheet.build_charts }

    it 'builds the correct number of charts' do
      expect(subject.size).to eq(0)
    end

  end

  describe '#build_width_normalization' do

    subject { default_sheet.build_width_normalization }

    it 'builds an empty hash' do
      expect(subject).to eq({})
    end

  end

end
