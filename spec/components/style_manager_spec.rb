require 'spec_helper'

describe Rubyxls::Components::StyleManager do
  let(:default_style_manager) { Rubyxls::Components::StyleManager.new }

  describe '#retrieve_style_attributes' do

    context 'single style attributes' do

      subject { default_style_manager.retrieve_style_attributes([:bold]) }

      it 'retrieves the correct attributes' do
        expect(subject).to eq(
                              {
                                alignment: { horizontal: :left, wrap_text: false, vertical: :bottom },
                                b: true,
                                font_name: "Arial",
                                border: { style: :none, color: "000000"},
                                sz: 10
                              }
                            )
      end

    end

    context 'combined style attributes' do

      subject { default_style_manager.retrieve_style_attributes([:bold, :italic]) }

      it 'retrieves the correct attributes' do
        expect(subject).to eq(
                              {
                                alignment: { horizontal: :left, wrap_text: false, vertical: :bottom },
                                b: true,
                                border: { style: :none, color: "000000"},
                                font_name: "Arial",
                                i: true,
                                sz: 10
                              }
                             )
      end

    end

    context 'style is not defined' do

      subject { default_style_manager.retrieve_style_attributes([:fake_style]) }

      it 'raises an error' do
        expect { subject }.to raise_error("The style :fake_style has not been defined! Please define or create a new type of Style in the styles directory!")
      end

    end

  end

  describe 'defined styles' do

    it 'has default style defined' do
      expect(default_style_manager.retrieve_style_attributes([:default])).to_not be_nil
    end

    it 'has base styles defined' do
      [:bold, :italic, :underline, :strike, :indent, :left_align, :right_align, :center_align, :top_align, :bottom_align, :middle_align, :border_right, :border_left, :border_top, :border_bottom, :border_all, :number, :decimal, :date, :time, :currency, :currency_precision, :percent, :wrap_text].each do |style|
        expect(default_style_manager.retrieve_style_attributes([style])).to_not be_nil
      end
    end

  end

end
