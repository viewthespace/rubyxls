require 'spec_helper'

describe Reporting::Excel2::ViewModels::DefaultViewModel do
  let(:default_view_model) { Reporting::Excel2::ViewModels::DefaultViewModel.new }

  describe '#rows' do

    subject { default_view_model.data_rows }

    it 'builds default' do
      expect(subject).to eq([])
    end

  end

end
