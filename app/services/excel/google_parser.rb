module Excel
  class GoogleParser < BaseParser
    # 구글 정산 파일
    # 특수 구조: 헤더행 이후부터 데이터 시작
    # 담당자 컬럼을 계정 식별자로 사용 (D열/인덱스 3)

    def account_id_column
      3 # D열 (담당자/계정명)
    end

    def parse
      open_workbook
      return { rows: [], errors: @errors } unless valid?

      sheet = @workbook.sheet(0)

      # 헤더 행 찾기 ("계정 ID" 텍스트가 있는 행)
      header_row = find_header_row(sheet)
      unless header_row
        @errors << "구글 정산 데이터 헤더를 찾을 수 없습니다"
        return { rows: [], errors: @errors }
      end

      headers = sheet.row(header_row)
      rows = []

      (header_row + 1..sheet.last_row).each do |row_num|
        row_data = sheet.row(row_num)
        next if row_data.compact.empty?

        # 첫 번째 컬럼(A)이 비어있고, 두 번째 컬럼(B)이 숫자인 경우만 데이터 행
        next unless row_data[0].nil? && row_data[1].is_a?(Numeric)

        # 합계 행 스킵 (B열이 nil이면)
        next if row_data[1].nil?

        account_id = extract_account_id(row_data)
        next if account_id.nil? || account_id.empty?

        rows << build_row_hash(row_num, row_data, headers, account_id)
      end

      { rows: rows, errors: @errors, total_rows: rows.size }
    end

    private

    def find_header_row(sheet)
      (1..20).each do |row_num|
        row = sheet.row(row_num)
        if row.any? { |cell| cell.to_s.include?("계정 ID") || cell.to_s.include?("계정ID") }
          return row_num
        end
      end
      nil
    end

    def extract_account_id(row_data)
      # D열 (인덱스 3)의 계정/담당자명을 계정 ID로 사용
      val = row_data[3]
      return nil if val.nil?
      # 구글은 한글 계정명을 사용하므로 단순 strip만 적용
      val.to_s.strip
    end

    def build_row_hash(row_num, row_data, headers, account_id)
      {
        row_number: row_num,
        account_id: account_id,
        sequence: row_data[1],             # B열: 순번
        account_type: row_data[2],         # C열: 계정 유형 (A-후불 등)
        account_name: row_data[3],         # D열: 계정/담당자
        supply_amount: row_data[4],        # E열: 공급가액
        da_supply_amount: row_data[5],     # F열: DA 공급가액
        raw_data: row_data,
        headers: headers
      }
    end
  end
end
