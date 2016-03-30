require 'spec_helper'

describe Rubyxls::ViewModel do
  let(:default_view_model) { Rubyxls::ViewModel.new }

  describe '#rows' do

    subject { default_view_model.data_rows }

    it 'builds default' do
      expect(subject).to eq([])
    end

  end

end
