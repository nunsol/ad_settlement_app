module Excel
  class NaverParser < BaseParser
    # 네이버 SA/GFA 정산 파일
    # 계정ID 위치: H열 (광고계정 AdAccountNo)

    ACCOUNT_ID_COLUMN = 8 # H열

    def account_id_column
      ACCOUNT_ID_COLUMN
    end

    def parse
      open_workbook
      return { rows: [], errors: @errors } unless valid?

      # 첫 번째 시트 (실적 데이터)
      sheet = @workbook.sheet(0)
      headers = sheet.row(1)

      rows = []
      (2..sheet.last_row).each do |row_num|
        row_data = sheet.row(row_num)
        next if row_data.compact.empty?

        account_id = extract_account_id(row_data)
        next if account_id.nil? || account_id.empty?

        rows << build_row_hash(row_num, row_data, headers, account_id)
      end

      { rows: rows, errors: @errors, total_rows: rows.size }
    end

    private

    def build_row_hash(row_num, row_data, headers, account_id)
      {
        row_number: row_num,
        account_id: account_id,
        agency: row_data[0],                # A열: 대행사
        agent_id: row_data[1],              # B열: 에이전트ID
        agent_name: row_data[2],            # C열: 담당 에이전트 이름
        customer_id: row_data[6],           # G열: 광고계정 CustomerID
        ad_account_no: row_data[7],         # H열: 광고계정 AdAccountNo (계정ID)
        advertiser_name: row_data[8],       # I열: 광고주명
        payment_type: row_data[9],          # J열: 결제유형
        raw_data: row_data,
        headers: headers
      }
    end
  end
end
