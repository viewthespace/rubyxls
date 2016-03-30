require 'spec_helper'

describe Rubyxls::Builders::RowBuilder do
  let(:single_column_letter_cells) {
                                      [
                                        { value: "cell1", row: 2, column: "B" }, { value: "cell2", row: 2, column: "C" },
                                        { value: "cell3", row: 3, column: "B" }, { value: "cell4", row: 3, column: "C" },
                                        { value: "cell5", row: 4, column: "B" }, { value: "cell6", row: 4, column: "C" },
                                        { value: "cell7", row: 5, column: "B" }, { value: "cell8", row: 5, column: "C" },
                                        { value: "cell9", row: 6, column: "B" }, { value: "cell10", row: 6, column: "C" }
                                      ]
                                    }
  let(:multi_column_letter_cells) {
                                    [
                                      { value: "cell1", row: 2, column: "AC" }, { value: "cell2", row: 2, column: "AD" },
                                      { value: "cell3", row: 3, column: "AC" }, { value: "cell4", row: 3, column: "AD" },
                                      { value: "cell5", row: 4, column: "AC" }, { value: "cell6", row: 4, column: "AD" },
                                      { value: "cell7", row: 5, column: "AC" }, { value: "cell8", row: 5, column: "AD" },
                                      { value: "cell9", row: 6, column: "AC" }, { value: "cell10", row: 6, column: "AD" }
                                    ]
                                  }

  describe '#rows' do

    context 'single column letter cells' do

      subject { Rubyxls::Builders::RowBuilder.new(single_column_letter_cells).rows }

      it 'maps each cell in specified column and row' do
        expect(subject).to eq(
                                [
                                  [{ value: nil, style: [] }, { value: nil, style: [] }, { value: nil, style: [] }],
                                  [{ value: nil, style: [] }, { value: "cell1", row: 2, column: "B" }, { value: "cell2", row: 2, column: "C"} ],
                                  [{ value: nil, style: [] }, { value: "cell3", row: 3, column: "B" }, { value: "cell4", row: 3, column: "C"} ],
                                  [{ value: nil, style: [] }, { value: "cell5", row: 4, column: "B" }, { value: "cell6", row: 4, column: "C"} ],
                                  [{ value: nil, style: [] }, { value: "cell7", row: 5, column: "B" }, { value: "cell8", row: 5, column: "C"} ],
                                  [{ value: nil, style: [] }, { value: "cell9", row: 6, column: "B" }, { value: "cell10", row: 6, column: "C"} ]
                                ]
                              )
      end

    end

    context 'multi column letter cells' do

      subject { Rubyxls::Builders::RowBuilder.new(multi_column_letter_cells).rows }

      it 'maps each cell in specified column and row' do
        expect(subject).to eq(
                                [
                                  [{ value: nil, style: [] }, { value: nil, style: [] }, { value: nil, style: [] }, { value: nil, style: [] }, { value: nil, style: [] }, { value: nil, style: [] }, { value: nil, style: [] }, { value: nil, style: [] }, { value: nil, style: [] }, { value: nil, style: [] }, { value: nil, style: [] }, { value: nil, style: [] }, { value: nil, style: [] }, { value: nil, style: [] }, { value: nil, style: [] }, { value: nil, style: [] }, { value: nil, style: [] }, { value: nil, style: [] }, { value: nil, style: [] }, { value: nil, style: [] }, { value: nil, style: [] }, { value: nil, style: [] }, { value: nil, style: [] }, { value: nil, style: [] }, { value: nil, style: [] }, { value: nil, style: [] }, { value: nil, style: [] }, { value: nil, style: [] }, { value: nil, style: [] }, { value: nil, style: [] }],
                                  [{ value: nil, style: [] }, { value: nil, style: [] }, { value: nil, style: [] }, { value: nil, style: [] }, { value: nil, style: [] }, { value: nil, style: [] }, { value: nil, style: [] }, { value: nil, style: [] }, { value: nil, style: [] }, { value: nil, style: [] }, { value: nil, style: [] }, { value: nil, style: [] }, { value: nil, style: [] }, { value: nil, style: [] }, { value: nil, style: [] }, { value: nil, style: [] }, { value: nil, style: [] }, { value: nil, style: [] }, { value: nil, style: [] }, { value: nil, style: [] }, { value: nil, style: [] }, { value: nil, style: [] }, { value: nil, style: [] }, { value: nil, style: [] }, { value: nil, style: [] }, { value: nil, style: [] }, { value: nil, style: [] }, { value: nil, style: [] }, { value: "cell1", row: 2, column: "AC" }, { value: "cell2", row: 2, column: "AD" }],
                                  [{ value: nil, style: [] }, { value: nil, style: [] }, { value: nil, style: [] }, { value: nil, style: [] }, { value: nil, style: [] }, { value: nil, style: [] }, { value: nil, style: [] }, { value: nil, style: [] }, { value: nil, style: [] }, { value: nil, style: [] }, { value: nil, style: [] }, { value: nil, style: [] }, { value: nil, style: [] }, { value: nil, style: [] }, { value: nil, style: [] }, { value: nil, style: [] }, { value: nil, style: [] }, { value: nil, style: [] }, { value: nil, style: [] }, { value: nil, style: [] }, { value: nil, style: [] }, { value: nil, style: [] }, { value: nil, style: [] }, { value: nil, style: [] }, { value: nil, style: [] }, { value: nil, style: [] }, { value: nil, style: [] }, { value: nil, style: [] }, { value: "cell3", row: 3, column: "AC" }, { value: "cell4", row: 3, column: "AD" }],
                                  [{ value: nil, style: [] }, { value: nil, style: [] }, { value: nil, style: [] }, { value: nil, style: [] }, { value: nil, style: [] }, { value: nil, style: [] }, { value: nil, style: [] }, { value: nil, style: [] }, { value: nil, style: [] }, { value: nil, style: [] }, { value: nil, style: [] }, { value: nil, style: [] }, { value: nil, style: [] }, { value: nil, style: [] }, { value: nil, style: [] }, { value: nil, style: [] }, { value: nil, style: [] }, { value: nil, style: [] }, { value: nil, style: [] }, { value: nil, style: [] }, { value: nil, style: [] }, { value: nil, style: [] }, { value: nil, style: [] }, { value: nil, style: [] }, { value: nil, style: [] }, { value: nil, style: [] }, { value: nil, style: [] }, { value: nil, style: [] }, { value: "cell5", row: 4, column: "AC" }, { value: "cell6", row: 4, column: "AD" }],
                                  [{ value: nil, style: [] }, { value: nil, style: [] }, { value: nil, style: [] }, { value: nil, style: [] }, { value: nil, style: [] }, { value: nil, style: [] }, { value: nil, style: [] }, { value: nil, style: [] }, { value: nil, style: [] }, { value: nil, style: [] }, { value: nil, style: [] }, { value: nil, style: [] }, { value: nil, style: [] }, { value: nil, style: [] }, { value: nil, style: [] }, { value: nil, style: [] }, { value: nil, style: [] }, { value: nil, style: [] }, { value: nil, style: [] }, { value: nil, style: [] }, { value: nil, style: [] }, { value: nil, style: [] }, { value: nil, style: [] }, { value: nil, style: [] }, { value: nil, style: [] }, { value: nil, style: [] }, { value: nil, style: [] }, { value: nil, style: [] }, { value: "cell7", row: 5, column: "AC" }, { value: "cell8", row: 5, column: "AD" }],
                                  [{ value: nil, style: [] }, { value: nil, style: [] }, { value: nil, style: [] }, { value: nil, style: [] }, { value: nil, style: [] }, { value: nil, style: [] }, { value: nil, style: [] }, { value: nil, style: [] }, { value: nil, style: [] }, { value: nil, style: [] }, { value: nil, style: [] }, { value: nil, style: [] }, { value: nil, style: [] }, { value: nil, style: [] }, { value: nil, style: [] }, { value: nil, style: [] }, { value: nil, style: [] }, { value: nil, style: [] }, { value: nil, style: [] }, { value: nil, style: [] }, { value: nil, style: [] }, { value: nil, style: [] }, { value: nil, style: [] }, { value: nil, style: [] }, { value: nil, style: [] }, { value: nil, style: [] }, { value: nil, style: [] }, { value: nil, style: [] }, { value: "cell9", row: 6, column: "AC" }, { value: "cell10", row: 6, column: "AD" }]
                                ]
                              )
      end

    end

  end

end
