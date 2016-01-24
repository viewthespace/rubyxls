require 'spec_helper'

describe Rubyxml::Components::ViewModel do
  let(:default_view_model) { Rubyxml::Components::ViewModel.new }

  describe '#rows' do

    subject { default_view_model.data_rows }

    it 'builds default' do
      expect(subject).to eq([])
    end

  end

end
