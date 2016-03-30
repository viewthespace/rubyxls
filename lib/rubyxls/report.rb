module Rubyxls
  class Report

    attr_reader :workbooks
    attr_reader :file_name

    def initialize
      @workbooks = []
      @file_name = 'rubyxls'
      build_workbooks!
    end

    def download!
      if @workbooks.size == 1
        @workbooks.first.to_stream
      else
        stream = StringIO.new
        zipfile = Zip::File.new(stream, true, true)
        @workbooks.each do |workbook|
          zipfile.get_output_stream(workbook.filename) { |zipstream| IO.copy_stream(workbook.to_stream, zipstream) }
        end
        zipfile.write_buffer(stream)
        zipfile.glob('*', &:clean_up)
        stream.rewind
        stream
      end
    end

    def file_extension
      @workbooks.size > 1 ? :zip : :xlsx
    end

    def content_type
      if file_extension == :xlsx
        'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
      elsif file_extension == :zip
        'application/zip'
      end
    end

    private

    def build_workbooks!
      @workbooks << Rubyxls::Workbook.new
    end

  end
end
