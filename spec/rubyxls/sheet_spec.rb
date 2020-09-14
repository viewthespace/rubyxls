require 'spec_helper'

describe Rubyxls::Sheet do
  let(:default_sheet) { Rubyxls::Sheet.new }

  describe '#build_options' do
    subject { default_sheet.build_options }

    it 'builds the correct options' do
      expect(subject).to eq({
                              page_setup: { fit_to_width: 1, orientation: :landscape, paper_size: 1 },
                              print_options: { grid_lines: false },
                              page_margins: { top: 0.3, left: 0.3, bottom: 0.3, right: 0.3, header: 0.0, footer: 0.0 },
                              name: "Default Sheet"
                            })
    end

    context 'dirty sheet name' do
      let(:dirty_sheet) { Rubyxls::Sheet.new(sheet_name: "T:[is] /s? D**t\\:") }

      subject { dirty_sheet.build_options }

      it 'sets clean sheet name' do
        expect(subject[:name]).to eq("T--is- -s- D--t--")
      end
    end

    describe 'sheet name' do
      context 'when name is longer than 31 bytesize' do
        let(:long_sheet) { Rubyxls::Sheet.new(sheet_name: "This sheet name is way too longgggggggggggg") }

        subject { long_sheet.build_options }

        it 'cuts the sheet name off at 31 bytesize chars' do
          expect(subject[:name].bytesize).to eq(31)
        end
      end

      context 'when sheet name has chars greater than 1 byte' do
        let(:long_sheet) { Rubyxls::Sheet.new(sheet_name: "1234567890 – 1234567890 1234567890 123") }

        subject { long_sheet.build_options }

        it 'cuts the sheet name off at 31 bytesize chars' do
          expect(subject[:name].bytesize).to eq(31)
        end
      end

      context 'when sheet name has chars greater than 1 byte and name is taken' do
        let(:long_sheet) { Rubyxls::Sheet.new(sheet_name: '1234567890 – 1234567890 1234567890 1') }

        subject { long_sheet.build_options('1234567890 – 1234567890 12345') }

        it 'cuts the sheet name off at 31 bytesize chars' do
          expect(subject[:name].bytesize).to eq(31)
          expect(subject[:name]).to eq('1234567890 – 1234567890 12(1)')
        end
      end
    end

    context 'taken sheet' do
      subject { default_sheet.build_options("Default Sheet") }

      it 'sets an available sheet name' do
        expect(subject[:name]).to eq("Default Sheet(1)")
      end
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
