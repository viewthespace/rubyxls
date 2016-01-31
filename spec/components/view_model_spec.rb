require 'spec_helper'

describe Rubyxls::Components::ViewModel do
  let(:default_view_model) { Rubyxls::Components::ViewModel.new }

  describe '#rows' do

    subject { default_view_model.data_rows }

    it 'builds default' do
      expect(subject).to eq([])
    end

  end

end
