require 'spec_helper'

describe Rubyxls::ViewModel do
  describe '#rows' do
    let(:default_view_model) { Rubyxls::ViewModel.new }

    subject { default_view_model.data_rows }

    it 'builds default' do
      expect(subject).to eq([])
    end

  end

  describe '#build_additional_rows!' do
    context 'when there are no other rows' do
      class NoOtherRowsViewModel < Rubyxls::ViewModel
        def build_additional_row!
          @data_rows << [{data: 'additional row'}]
        end
      end

      let(:no_other_rows_view_model) { NoOtherRowsViewModel.new(rows_count: 4) }

      it 'fills data rows with the additional rows' do
        expect(no_other_rows_view_model.data_rows).to eq(
          [
            [{data: 'additional row'}],
            [{data: 'additional row'}],
            [{data: 'additional row'}],
            [{data: 'additional row'}]
          ]
        )
      end
    end

    context 'when there are some other rows but more than the specified amount' do
      class SomeOtherRowsViewModel < Rubyxls::ViewModel
        def build_title_row!
          @data_rows << [{data: 'title row'}]
        end

        def build_header_row!
          @data_rows << [{data: 'header row'}]
        end

        def build_additional_row!
          @data_rows << [{data: 'additional row'}]
        end
      end

      let(:some_other_rows_view_model) { SomeOtherRowsViewModel.new(rows_count: 4) }

      it 'fills data rows with the additional rows' do
        expect(some_other_rows_view_model.data_rows).to eq(
          [
            [{data: 'title row'}],
            [{data: 'header row'}],
            [{data: 'additional row'}],
            [{data: 'additional row'}]
          ]
        )
      end
    end

    context 'when there are more rows then specified amount' do
      class OverflowRowsViewModel < Rubyxls::ViewModel
        def build_title_row!
          @data_rows << [{data: 'title row'}]
        end

        def build_header_row!
          @data_rows << [{data: 'header row'}]
        end

        def build_data_rows!
          ['data1', 'data2', 'data3'].each do |data|
            @data_rows << [{data: data}]
          end
        end

        def build_additional_row!
          @data_rows << [{data: 'additional row'}]
        end
      end

      let(:overflow_rows_view_model) { OverflowRowsViewModel.new(rows_count: 4) }

      it 'adds no additional rows' do
        expect(overflow_rows_view_model.data_rows).to eq(
          [
            [{data: 'title row'}],
            [{data: 'header row'}],
            [{data: 'data1'}],
            [{data: 'data2'}],
            [{data: 'data3'}]
          ]
        )
      end
    end
  end
end
