require 'spec_helper'

describe Rubyxls do

  it 'has correct version number' do
    expect(Rubyxls::VERSION).to eq("0.1.0")
  end

  describe '#generate_default_report' do
    subject { Rubyxls.generate_default_report }

    it 'generates default instance of Report' do
      expect(subject).to be_an_instance_of(Rubyxls::Report)
    end
  end

  describe '#write_report_to_tmp' do
    let(:report) { Rubyxls.generate_default_report }

    before do
      Rubyxls.write_report_to_tmp(report)
    end

    it 'writes report to current dir' do
      expect(File.exist?(file_path_to_tmp(report))).to eq(true)
    end
  end

  def file_path_to_tmp(report)
    File.expand_path('../../tmp', __FILE__) + "/#{report.file_name}.#{report.file_extension}"
  end

end
