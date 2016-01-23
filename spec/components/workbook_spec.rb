require 'spec_helper'

describe Reporting::Excel2::Workbooks::DefaultWorkbook do
  let(:workbook) { Reporting::Excel2::Workbooks::DefaultWorkbook.new }

  describe '#to_stream' do

    subject { workbook.to_stream }

    it 'returns an instance of StringIO' do
      expect(subject).to be_a(StringIO)
    end

  end

  describe '#sheets' do

    subject { workbook.sheets.size }

    it 'returns the correct number of sheets' do
      expect(subject).to eq(1)
    end

  end

  describe '#filename' do

    subject { workbook.filename }

    it 'is correct' do
      expect(subject).to eq("default_workbook.xlsx")
    end

  end
  
end