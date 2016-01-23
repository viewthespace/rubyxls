require 'spec_helper'

describe Reporting::Excel2::Builders::FormulaBuilder do

  describe '#cells' do

    context 'fill right' do

      context 'no locked rows and no locked columns' do

        subject { Reporting::Excel2::Builders::FormulaBuilder.new( value: "=B1+B2", style: [:bold], width: 10, height: 10, fill: :right, num_cells: 4).cells }

        it 'has the correct cells' do
          expect(subject).to eq([{ value: "=B1+B2", style: [:bold], width: 10, height: 10, data_validation: nil, merge: nil }, { value: "=C1+C2", style: [:bold], width: 10, height: 10, data_validation: nil, merge: nil }, { value: "=D1+D2", style: [:bold], width: 10, height: 10, data_validation: nil, merge: nil }, { value: "=E1+E2", style: [:bold], width: 10, height: 10, data_validation: nil, merge: nil }])
        end

      end

      context 'locked rows' do

        subject { Reporting::Excel2::Builders::FormulaBuilder.new( value: "=B$1+B$2", style: [:bold], width: 10, height: 10, fill: :right, num_cells: 4).cells }

        it 'has the correct cells' do
          expect(subject).to eq([{ value: "=B$1+B$2", style: [:bold], width: 10, height: 10, data_validation: nil, merge: nil }, { value: "=C$1+C$2", style: [:bold], width: 10, height: 10, data_validation: nil, merge: nil }, { value: "=D$1+D$2", style: [:bold], width: 10, height: 10, data_validation: nil, merge: nil }, { value: "=E$1+E$2", style: [:bold], width: 10, height: 10, data_validation: nil, merge: nil }])
        end

      end

      context 'locked columns' do

        subject { Reporting::Excel2::Builders::FormulaBuilder.new( value: "=$B1+$B2", style: [:bold], width: 10, height: 10, fill: :right, num_cells: 4).cells }

        it 'has the correct cells' do
          expect(subject).to eq([{ value: "=$B1+$B2", style: [:bold], width: 10, height: 10, data_validation: nil, merge: nil }, { value: "=$B1+$B2", style: [:bold], width: 10, height: 10, data_validation: nil, merge: nil }, { value: "=$B1+$B2", style: [:bold], width: 10, height: 10, data_validation: nil, merge: nil }, { value: "=$B1+$B2", style: [:bold], width: 10, height: 10, data_validation: nil, merge: nil }])
        end

      end

      context 'locked rows and locked columns' do

        subject { Reporting::Excel2::Builders::FormulaBuilder.new( value: "=$B$1+$B$2", style: [:bold], width: 10, height: 10, fill: :right, num_cells: 4).cells }

        it 'has the correct cells' do
          expect(subject).to eq([{ value: "=$B$1+$B$2", style: [:bold], width: 10, height: 10, data_validation: nil, merge: nil }, { value: "=$B$1+$B$2", style: [:bold], width: 10, height: 10, data_validation: nil, merge: nil }, { value: "=$B$1+$B$2", style: [:bold], width: 10, height: 10, data_validation: nil, merge: nil }, { value: "=$B$1+$B$2", style: [:bold], width: 10, height: 10, data_validation: nil, merge: nil }])
        end
      
      end

      context 'num_cells is nil' do

        subject { Reporting::Excel2::Builders::FormulaBuilder.new( value: "=B1+B2", style: [:bold], width: 10, height: 10, fill: :right).cells }

        it 'raises an error' do
          expect { subject }.to raise_error("Num cells cannot be nil when filling right for formula: =B1+B2!")
        end

      end

    end

    context 'fill down' do

      context 'no locked rows and no locked columns' do

        subject { 4.times.map { |n| Reporting::Excel2::Builders::FormulaBuilder.new( value: "=B1+B2", style: [:bold], width: 10, height: 10, fill: :down, row_index: n).cells } }

        it 'has the correct cells' do
          expect(subject).to eq([[{ value: "=B1+B2", style: [:bold], width: 10, height: 10, data_validation: nil, merge: nil }], [{ value: "=B2+B3", style: [:bold], width: 10, height: 10, data_validation: nil, merge: nil }], [{ value: "=B3+B4", style: [:bold], width: 10, height: 10, data_validation: nil, merge: nil }], [{ value: "=B4+B5", style: [:bold], width: 10, height: 10, data_validation: nil, merge: nil }]])
        end

      end

      context 'locked rows' do

        subject { 4.times.map { |n| Reporting::Excel2::Builders::FormulaBuilder.new( value: "=B$1+B$2", style: [:bold], width: 10, height: 10, fill: :down, row_index: n).cells } }

        it 'has the correct cells' do
          expect(subject).to eq([[{ value: "=B$1+B$2", style: [:bold], width: 10, height: 10, data_validation: nil, merge: nil }], [{ value: "=B$1+B$2", style: [:bold], width: 10, height: 10, data_validation: nil, merge: nil }], [{ value: "=B$1+B$2", style: [:bold], width: 10, height: 10, data_validation: nil, merge: nil }], [{ value: "=B$1+B$2", style: [:bold], width: 10, height: 10, data_validation: nil, merge: nil }]])
        end

      end

      context 'locked columns' do

        subject { 4.times.map { |n| Reporting::Excel2::Builders::FormulaBuilder.new( value: "=$B1+$B2", style: [:bold], width: 10, height: 10, fill: :down, row_index: n).cells } }

        it 'has the correct cells' do
          expect(subject).to eq([[{ value: "=$B1+$B2", style: [:bold], width: 10, height: 10, data_validation: nil, merge: nil }], [{ value: "=$B2+$B3", style: [:bold], width: 10, height: 10, data_validation: nil, merge: nil }], [{ value: "=$B3+$B4", style: [:bold], width: 10, height: 10, data_validation: nil, merge: nil }], [{ value: "=$B4+$B5", style: [:bold], width: 10, height: 10, data_validation: nil, merge: nil }]])
        end
        
      end

      context 'locked rows and locked columns' do

        subject { 4.times.map { |n| Reporting::Excel2::Builders::FormulaBuilder.new( value: "=$B$1+$B$2", style: [:bold], width: 10, height: 10, fill: :down, row_index: n).cells } }

        it 'has the correct cells' do
          expect(subject).to eq([[{ value: "=$B$1+$B$2", style: [:bold], width: 10, height: 10, data_validation: nil, merge: nil }], [{ value: "=$B$1+$B$2", style: [:bold], width: 10, height: 10, data_validation: nil, merge: nil }], [{ value: "=$B$1+$B$2", style: [:bold], width: 10, height: 10, data_validation: nil, merge: nil }], [{ value: "=$B$1+$B$2", style: [:bold], width: 10, height: 10, data_validation: nil, merge: nil }]])
        end

      end

      context 'row_index is nil' do

        subject { 4.times.map { |n| Reporting::Excel2::Builders::FormulaBuilder.new( value: "=B1+B2", style: [:bold], width: 10, height: 10, fill: :down).cells } }

        it 'raises an error' do
          expect { subject }.to raise_error("Row index cannot be nil when filling down for formula: =B1+B2!")
        end

      end

    end

    context 'fill all' do

      context 'no locked rows and no locked columns' do

        subject { 4.times.map { |n| Reporting::Excel2::Builders::FormulaBuilder.new( value: "=B1+B2", style: [:bold], width: 10, height: 10, fill: :all, num_cells: 4, row_index: n).cells } }

        it 'has the correct cells' do
          expect(subject).to eq([
                                  [{ value: "=B1+B2", style: [:bold], width: 10, height: 10, data_validation: nil, merge: nil }, { value: "=C1+C2", style: [:bold], width: 10, height: 10, data_validation: nil, merge: nil }, { value: "=D1+D2", style: [:bold], width: 10, height: 10, data_validation: nil, merge: nil }, { value: "=E1+E2", style: [:bold], width: 10, height: 10, data_validation: nil, merge: nil }],
                                  [{ value: "=B2+B3", style: [:bold], width: 10, height: 10, data_validation: nil, merge: nil }, { value: "=C2+C3", style: [:bold], width: 10, height: 10, data_validation: nil, merge: nil }, { value: "=D2+D3", style: [:bold], width: 10, height: 10, data_validation: nil, merge: nil }, { value: "=E2+E3", style: [:bold], width: 10, height: 10, data_validation: nil, merge: nil }],
                                  [{ value: "=B3+B4", style: [:bold], width: 10, height: 10, data_validation: nil, merge: nil }, { value: "=C3+C4", style: [:bold], width: 10, height: 10, data_validation: nil, merge: nil }, { value: "=D3+D4", style: [:bold], width: 10, height: 10, data_validation: nil, merge: nil }, { value: "=E3+E4", style: [:bold], width: 10, height: 10, data_validation: nil, merge: nil }],
                                  [{ value: "=B4+B5", style: [:bold], width: 10, height: 10, data_validation: nil, merge: nil }, { value: "=C4+C5", style: [:bold], width: 10, height: 10, data_validation: nil, merge: nil }, { value: "=D4+D5", style: [:bold], width: 10, height: 10, data_validation: nil, merge: nil }, { value: "=E4+E5", style: [:bold], width: 10, height: 10, data_validation: nil, merge: nil }]
                               ])
        end

      end

      context 'locked rows' do

        subject { 4.times.map { |n| Reporting::Excel2::Builders::FormulaBuilder.new( value: "=B$1+B$2", style: [:bold], width: 10, height: 10, fill: :all, num_cells: 4, row_index: n).cells } }

        it 'has the correct cells' do
          expect(subject).to eq([
                                  [{ value: "=B$1+B$2", style: [:bold], width: 10, height: 10, data_validation: nil, merge: nil }, { value: "=C$1+C$2", style: [:bold], width: 10, height: 10, data_validation: nil, merge: nil }, { value: "=D$1+D$2", style: [:bold], width: 10, height: 10, data_validation: nil, merge: nil }, { value: "=E$1+E$2", style: [:bold], width: 10, height: 10, data_validation: nil, merge: nil }],
                                  [{ value: "=B$1+B$2", style: [:bold], width: 10, height: 10, data_validation: nil, merge: nil }, { value: "=C$1+C$2", style: [:bold], width: 10, height: 10, data_validation: nil, merge: nil }, { value: "=D$1+D$2", style: [:bold], width: 10, height: 10, data_validation: nil, merge: nil }, { value: "=E$1+E$2", style: [:bold], width: 10, height: 10, data_validation: nil, merge: nil }],
                                  [{ value: "=B$1+B$2", style: [:bold], width: 10, height: 10, data_validation: nil, merge: nil }, { value: "=C$1+C$2", style: [:bold], width: 10, height: 10, data_validation: nil, merge: nil }, { value: "=D$1+D$2", style: [:bold], width: 10, height: 10, data_validation: nil, merge: nil }, { value: "=E$1+E$2", style: [:bold], width: 10, height: 10, data_validation: nil, merge: nil }],
                                  [{ value: "=B$1+B$2", style: [:bold], width: 10, height: 10, data_validation: nil, merge: nil }, { value: "=C$1+C$2", style: [:bold], width: 10, height: 10, data_validation: nil, merge: nil }, { value: "=D$1+D$2", style: [:bold], width: 10, height: 10, data_validation: nil, merge: nil }, { value: "=E$1+E$2", style: [:bold], width: 10, height: 10, data_validation: nil, merge: nil }]
                               ])
        end

      end

      context 'locked columns' do

        subject { 4.times.map { |n| Reporting::Excel2::Builders::FormulaBuilder.new( value: "=$B1+$B2", style: [:bold], width: 10, height: 10, fill: :all, num_cells: 4, row_index: n).cells } }

        it 'has the correct cells' do
          expect(subject).to eq([
                                  [{ value: "=$B1+$B2", style: [:bold], width: 10, height: 10, data_validation: nil, merge: nil }, { value: "=$B1+$B2", style: [:bold], width: 10, height: 10, data_validation: nil, merge: nil }, { value: "=$B1+$B2", style: [:bold], width: 10, height: 10, data_validation: nil, merge: nil }, { value: "=$B1+$B2", style: [:bold], width: 10, height: 10, data_validation: nil, merge: nil }],
                                  [{ value: "=$B2+$B3", style: [:bold], width: 10, height: 10, data_validation: nil, merge: nil }, { value: "=$B2+$B3", style: [:bold], width: 10, height: 10, data_validation: nil, merge: nil }, { value: "=$B2+$B3", style: [:bold], width: 10, height: 10, data_validation: nil, merge: nil }, { value: "=$B2+$B3", style: [:bold], width: 10, height: 10, data_validation: nil, merge: nil }],
                                  [{ value: "=$B3+$B4", style: [:bold], width: 10, height: 10, data_validation: nil, merge: nil }, { value: "=$B3+$B4", style: [:bold], width: 10, height: 10, data_validation: nil, merge: nil }, { value: "=$B3+$B4", style: [:bold], width: 10, height: 10, data_validation: nil, merge: nil }, { value: "=$B3+$B4", style: [:bold], width: 10, height: 10, data_validation: nil, merge: nil }],
                                  [{ value: "=$B4+$B5", style: [:bold], width: 10, height: 10, data_validation: nil, merge: nil }, { value: "=$B4+$B5", style: [:bold], width: 10, height: 10, data_validation: nil, merge: nil }, { value: "=$B4+$B5", style: [:bold], width: 10, height: 10, data_validation: nil, merge: nil }, { value: "=$B4+$B5", style: [:bold], width: 10, height: 10, data_validation: nil, merge: nil }]
                               ])
        end
        
      end

      context 'locked rows and locked columns' do

        subject { 4.times.map { |n| Reporting::Excel2::Builders::FormulaBuilder.new( value: "=$B$1+$B$2", style: [:bold], width: 10, height: 10, fill: :all, num_cells: 4, row_index: 0).cells } }

        it 'has the correct cells' do
          expect(subject).to eq([
                                  [{ value: "=$B$1+$B$2", style: [:bold], width: 10, height: 10, data_validation: nil, merge: nil }, { value: "=$B$1+$B$2", style: [:bold], width: 10, height: 10, data_validation: nil, merge: nil }, { value: "=$B$1+$B$2", style: [:bold], width: 10, height: 10, data_validation: nil, merge: nil }, { value: "=$B$1+$B$2", style: [:bold], width: 10, height: 10, data_validation: nil, merge: nil }],
                                  [{ value: "=$B$1+$B$2", style: [:bold], width: 10, height: 10, data_validation: nil, merge: nil }, { value: "=$B$1+$B$2", style: [:bold], width: 10, height: 10, data_validation: nil, merge: nil }, { value: "=$B$1+$B$2", style: [:bold], width: 10, height: 10, data_validation: nil, merge: nil }, { value: "=$B$1+$B$2", style: [:bold], width: 10, height: 10, data_validation: nil, merge: nil }],
                                  [{ value: "=$B$1+$B$2", style: [:bold], width: 10, height: 10, data_validation: nil, merge: nil }, { value: "=$B$1+$B$2", style: [:bold], width: 10, height: 10, data_validation: nil, merge: nil }, { value: "=$B$1+$B$2", style: [:bold], width: 10, height: 10, data_validation: nil, merge: nil }, { value: "=$B$1+$B$2", style: [:bold], width: 10, height: 10, data_validation: nil, merge: nil }],
                                  [{ value: "=$B$1+$B$2", style: [:bold], width: 10, height: 10, data_validation: nil, merge: nil }, { value: "=$B$1+$B$2", style: [:bold], width: 10, height: 10, data_validation: nil, merge: nil }, { value: "=$B$1+$B$2", style: [:bold], width: 10, height: 10, data_validation: nil, merge: nil }, { value: "=$B$1+$B$2", style: [:bold], width: 10, height: 10, data_validation: nil, merge: nil }]
                               ])
        end

      end

      context 'num_cells is nil' do

        subject { 4.times.map { |n| Reporting::Excel2::Builders::FormulaBuilder.new( value: "=B1+B2", style: [:bold], width: 10, height: 10, fill: :all, row_index: n).cells } }

        it 'raises an error' do
          expect { subject }.to raise_error("Both num cells & row index cannot be nil when filling all for formula: =B1+B2!")
        end

      end

      context 'row_index is nil' do

        subject { 4.times.map { |n| Reporting::Excel2::Builders::FormulaBuilder.new( value: "=B1+B2", style: [:bold], width: 10, height: 10, fill: :all, num_cells: n).cells } }

        it 'raises an error' do
          expect { subject }.to raise_error("Both num cells & row index cannot be nil when filling all for formula: =B1+B2!")
        end

      end

    end

  end

end