require 'spec_helper'

describe Reporting::Excel2::Builders::CellBuilder do
  let(:model_data_rows) {
                     [
                       [{ value: "", style: [:default] }, { value: "", style: [:default] }, { value: "", style: [:default] }, { value: "", style: [:default] }],
                       [{ value: "", style: [:default] }, { value: "", style: [:default] }, { value: "", style: [:default] }, { value: "", style: [:default] }],
                       [{ value: "", style: [:default] }, { value: "", style: [:default] }, { value: "", style: [:default] }, { value: "", style: [:default] }],
                       [{ value: "", style: [:default] }, { value: "", style: [:default] }, { value: "", style: [:default] }, { value: "", style: [:default] }]
                     ]
                    }

  describe '#cells' do

    describe 'row numbers' do

      subject { Reporting::Excel2::Builders::CellBuilder.new(model_data_rows: model_data_rows, start_row: 2, start_column: "B").cells.map { |cell| cell[:row] } }

      it 'has the correct row number' do
        expect(subject).to eq([2, 2, 2, 2, 3, 3, 3, 3, 4, 4, 4, 4, 5, 5, 5, 5])
      end

    end

    describe 'column letters' do

      context 'single letter' do

        subject { Reporting::Excel2::Builders::CellBuilder.new(model_data_rows: model_data_rows, start_row: 2, start_column: "B").cells.map { |cell| cell[:column] } }

        it 'has the correct column letter' do
          expect(subject).to eq(["B", "C", "D", "E", "B", "C", "D", "E", "B", "C", "D", "E", "B", "C", "D", "E"])
        end

      end

      context 'multi letter' do

        subject { Reporting::Excel2::Builders::CellBuilder.new(model_data_rows: model_data_rows, start_row: 2, start_column: "AZ").cells.map { |cell| cell[:column] } }

        it 'has the correct column letters' do
          expect(subject).to eq(["AZ", "BA", "BB", "BC", "AZ", "BA", "BB", "BC", "AZ", "BA", "BB", "BC", "AZ", "BA", "BB", "BC"])
        end

      end

    end

  end

end