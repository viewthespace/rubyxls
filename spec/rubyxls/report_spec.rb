require 'spec_helper'

describe Rubyxls::Report do
  let(:default_report) { Rubyxls::Report.new }

  describe '#download!' do

    subject { default_report.download! }

    it 'returns an instance of StringIO' do
      expect(subject).to be_a(StringIO)
    end

  end

  describe '#workbooks' do

    subject { default_report.workbooks.size }

    it 'returns the correct number of workbooks' do
      expect(subject).to eq(1)
    end

  end

  describe '#file_extension' do

    subject { default_report.file_extension }

    it 'is correct' do
      expect(subject).to eq(:xlsx)
    end

  end

  describe '#content_type' do

    subject { default_report.content_type }

    it 'is correct' do
      expect(subject).to eq('application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
    end

  end

end
