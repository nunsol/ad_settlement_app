module Excel
  class BaseParser
    attr_reader :file_path, :workbook, :errors

    def initialize(file_path)
      @file_path = file_path
      @workbook = nil
      @errors = []
    end

    def parse
      raise NotImplementedError, "Subclasses must implement #parse"
    end

    def account_id_column
      raise NotImplementedError, "Subclasses must implement #account_id_column"
    end

    def valid?
      @errors.empty?
    end

    protected

    def open_workbook
      @workbook = Roo::Spreadsheet.open(@file_path)
    rescue StandardError => e
      @errors << "Failed to open file: #{e.message}"
      nil
    end

    def normalize_account_id(value)
      return nil if value.nil?
      value.to_s.strip.gsub(/[^0-9a-zA-Z\-_]/, "")
    end

    def extract_rows_with_account_ids(sheet_name: nil, header_row: 1, data_start_row: 2)
      open_workbook unless @workbook

      return [] unless @workbook

      sheet = sheet_name ? @workbook.sheet(sheet_name) : @workbook.sheet(0)
      headers = sheet.row(header_row)

      rows = []
      (data_start_row..sheet.last_row).each do |row_num|
        row_data = sheet.row(row_num)
        next if row_data.compact.empty?

        account_id = extract_account_id(row_data)
        next if account_id.nil? || account_id.empty?

        rows << {
          row_number: row_num,
          account_id: account_id,
          data: row_data,
          headers: headers
        }
      end

      rows
    end

    def extract_account_id(row_data)
      col_index = account_id_column - 1
      normalize_account_id(row_data[col_index])
    end
  end
end
