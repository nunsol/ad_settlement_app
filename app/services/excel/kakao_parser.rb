module Excel
  class KakaoParser < BaseParser
    # 카카오 모먼트/키워드/브랜드/채널 정산 파일
    # 계정ID 위치: C열 (자산 ID)
    # 동일한 포맷을 공유함

    ACCOUNT_ID_COLUMN = 3 # C열 (자산 ID)

    attr_reader :kakao_type

    def initialize(file_path, kakao_type: :moment)
      super(file_path)
      @kakao_type = kakao_type
    end

    def account_id_column
      ACCOUNT_ID_COLUMN
    end

    def parse
      open_workbook
      return { rows: [], errors: @errors } unless valid?

      # 첫 번째 시트 (정산 데이터)
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

      { rows: rows, errors: @errors, total_rows: rows.size, kakao_type: @kakao_type }
    end

    private

    def build_row_hash(row_num, row_data, headers, account_id)
      {
        row_number: row_num,
        account_id: account_id,
        kakao_type: @kakao_type,
        sales_date: row_data[0],           # A열: 매출일
        service: row_data[1],              # B열: 서비스
        asset_id: row_data[2],             # C열: 자산 ID (계정ID)
        asset_name: row_data[3],           # D열: 자산 이름
        advertiser_name: row_data[4],      # E열: 광고주 사업자명
        marketer_name: row_data[5],        # F열: 마케터명
        wallet_id: row_data[6],            # G열: 월렛 ID
        wallet_name: row_data[7],          # H열: 월렛 이름
        type: row_data[8],                 # I열: 유형
        business_number: row_data[9],      # J열: 사업자번호
        business_name: row_data[10],       # K열: 사업자명
        supply_amount: row_data[11],       # L열: 공급가액
        vat: row_data[12],                 # M열: 부가세
        free_cash: row_data[13],           # N열: 무상캐시
        raw_data: row_data,
        headers: headers
      }
    end
  end
end
