module Excel
  class TemplateProcessor
    # 매체별 통합 시트 이름
    MEDIA_SHEETS = {
      naver: "네이버통합_SA,GFA",
      kakao_moment: "카카오모먼트",
      kakao_keyword: "다음_검색(신)",
      kakao_brand: "다음_브랜드검색",
      kakao_channel: "다음_카카오톡채널",
      google: "구글"
    }.freeze

    # 대행사 시트 데이터 시작 행 (0-indexed, 행 18 = 인덱스 17)
    AGENCY_SHEET_DATA_START_ROW = 17

    # 대행사 시트 컬럼 인덱스 (0-indexed)
    AGENCY_COLUMNS = {
      media: 1,           # B열: 매체
      advertiser_name: 2, # C열: 광고주 명
      account_id: 3,      # D열: 광고주 ID
      supply_amount: 4,   # E열: 공급가액
      vat: 5,             # F열: VAT
      total: 6,           # G열: TOTAL
      fee_rate: 7,        # H열: 지급수수료율
      fee_supply: 8,      # I열: 수수료공급가
      fee_vat: 9,         # J열: 수수료부가세
      fee_total: 10       # K열: 수수료 합계
    }.freeze

    attr_reader :workbook, :errors, :distribution_stats

    def initialize(template_path, output_path)
      @template_path = template_path
      @output_path = output_path
      @workbook = nil
      @errors = []
      @distribution_stats = Hash.new { |h, k| h[k] = { rows: 0, accounts: [] } }
      @agency_row_counters = {}
    end

    def load!
      @workbook = RubyXL::Parser.parse(@template_path)
      initialize_agency_row_counters
      true
    rescue StandardError => e
      @errors << "템플릿 로드 실패: #{e.message}"
      false
    end

    # 대행사 시트에 정리된 데이터 추가
    def append_to_agency_sheet(agency_name, parsed_row, account_info:, source_type:)
      sheet = find_sheet(agency_name)

      unless sheet
        @errors << "시트를 찾을 수 없음: #{agency_name}"
        return false
      end

      # 현재 행 번호 가져오기
      row_num = next_agency_row(agency_name)

      # 데이터 변환 및 쓰기
      formatted_data = format_agency_row(parsed_row, account_info, source_type)
      write_agency_row(sheet, row_num, formatted_data)

      # 통계 업데이트
      @distribution_stats[agency_name][:rows] += 1
      @distribution_stats[agency_name][:accounts] << formatted_data[:account_id]

      true
    end

    # 매체 통합 시트에 원본 데이터 추가
    def append_to_media_sheet(media_type, row_data)
      sheet_name = MEDIA_SHEETS[media_type.to_sym]
      return false unless sheet_name

      sheet = find_sheet(sheet_name)
      return false unless sheet

      # 매체 시트는 기존 마지막 행 다음에 추가
      row_num = find_last_data_row(sheet) + 1

      row_data.each_with_index do |cell_value, col_idx|
        sheet.add_cell(row_num, col_idx, cell_value)
      end

      true
    end

    def save!
      @workbook.write(@output_path)
      true
    rescue StandardError => e
      @errors << "파일 저장 실패: #{e.message}"
      false
    end

    def sheet_names
      @workbook.worksheets.map(&:sheet_name)
    end

    def sheet_exists?(name)
      @workbook.worksheets.any? { |ws| ws.sheet_name == name }
    end

    private

    def find_sheet(name)
      @workbook.worksheets.find { |ws| ws.sheet_name == name }
    end

    def initialize_agency_row_counters
      # 대행사 시트는 데이터 시작 행부터 카운트
      @workbook.worksheets.each do |sheet|
        name = sheet.sheet_name
        # 기존 데이터가 있는지 확인하고 그 다음 행부터 시작
        existing_last_row = find_agency_last_data_row(sheet)
        @agency_row_counters[name] = [existing_last_row + 1, AGENCY_SHEET_DATA_START_ROW].max
      end
    end

    def find_agency_last_data_row(sheet)
      return AGENCY_SHEET_DATA_START_ROW - 1 if sheet.sheet_data.nil?

      last_row = AGENCY_SHEET_DATA_START_ROW - 1
      (AGENCY_SHEET_DATA_START_ROW..sheet.sheet_data.rows.size).each do |row_idx|
        row = sheet.sheet_data.rows[row_idx]
        next if row.nil?

        # B열(인덱스 1)에 데이터가 있으면 데이터 행으로 판단
        if row.cells && row.cells[1]&.value.present?
          last_row = row_idx
        end
      end
      last_row
    end

    def find_last_data_row(sheet)
      return 0 if sheet.sheet_data.nil?

      last_row = 0
      sheet.sheet_data.rows.each_with_index do |row, idx|
        next if row.nil?
        last_row = idx if row.cells&.any? { |c| c&.value.present? }
      end
      last_row
    end

    def next_agency_row(agency_name)
      @agency_row_counters[agency_name] ||= AGENCY_SHEET_DATA_START_ROW
      row = @agency_row_counters[agency_name]
      @agency_row_counters[agency_name] += 1
      row
    end

    # 매체별 raw 데이터를 대행사 시트 양식으로 변환
    def format_agency_row(parsed_row, account_info, source_type)
      raw = parsed_row[:raw_data] || []

      case source_type.to_s
      when "naver"
        format_naver_row(raw, account_info)
      when "kakao_moment", "kakao_keyword", "kakao_brand", "kakao_channel"
        format_kakao_row(raw, account_info, source_type)
      when "google"
        format_google_row(raw, account_info)
      else
        format_generic_row(parsed_row, account_info, source_type)
      end
    end

    def format_naver_row(raw, account_info)
      # 네이버 raw: H열(7)=계정ID, I열(8)=광고주명, 실적 컬럼들
      # 유상실적TOTAL은 M열(12)
      supply_amount = to_number(raw[12]) # 유상실적TOTAL
      vat = (supply_amount * 0.1).round
      total = supply_amount + vat
      fee_rate = account_info[:commission_rate] || 0.1
      fee_supply = (supply_amount * fee_rate).round
      fee_vat = (fee_supply * 0.1).round
      fee_total = fee_supply + fee_vat

      {
        media: "네이버",
        advertiser_name: raw[8]&.to_s || account_info[:advertiser_name],
        account_id: raw[7]&.to_s,
        supply_amount: supply_amount,
        vat: vat,
        total: total,
        fee_rate: fee_rate,
        fee_supply: fee_supply,
        fee_vat: fee_vat,
        fee_total: fee_total
      }
    end

    def format_kakao_row(raw, account_info, source_type)
      # 카카오 raw: C열(2)=자산ID, E열(4)=광고주명, L열(11)=공급가액, M열(12)=부가세
      media_name = case source_type.to_s
                   when "kakao_moment" then "카카오모먼트"
                   when "kakao_keyword" then "카카오키워드"
                   when "kakao_brand" then "카카오브랜드"
                   when "kakao_channel" then "카카오채널"
                   else "카카오"
                   end

      supply_amount = to_number(raw[11])
      vat = to_number(raw[12])
      total = supply_amount + vat
      fee_rate = account_info[:commission_rate] || 0.1
      fee_supply = (supply_amount * fee_rate).round
      fee_vat = (fee_supply * 0.1).round
      fee_total = fee_supply + fee_vat

      {
        media: media_name,
        advertiser_name: raw[4]&.to_s || account_info[:advertiser_name],
        account_id: raw[2]&.to_s,
        supply_amount: supply_amount,
        vat: vat,
        total: total,
        fee_rate: fee_rate,
        fee_supply: fee_supply,
        fee_vat: fee_vat,
        fee_total: fee_total
      }
    end

    def format_google_row(raw, account_info)
      # 구글 raw: D열(3)=계정ID, E열(4)=공급가액, F열(5)=DA공급가액
      supply_amount = to_number(raw[4])
      vat = (supply_amount * 0.1).round
      total = supply_amount + vat
      fee_rate = account_info[:commission_rate] || 0.03
      fee_supply = (supply_amount * fee_rate).round
      fee_vat = (fee_supply * 0.1).round
      fee_total = fee_supply + fee_vat

      {
        media: "구글",
        advertiser_name: account_info[:advertiser_name] || raw[3]&.to_s,
        account_id: raw[3]&.to_s,
        supply_amount: supply_amount,
        vat: vat,
        total: total,
        fee_rate: fee_rate,
        fee_supply: fee_supply,
        fee_vat: fee_vat,
        fee_total: fee_total
      }
    end

    def format_generic_row(parsed_row, account_info, source_type)
      {
        media: source_type.to_s,
        advertiser_name: account_info[:advertiser_name],
        account_id: parsed_row[:account_id],
        supply_amount: 0,
        vat: 0,
        total: 0,
        fee_rate: account_info[:commission_rate] || 0,
        fee_supply: 0,
        fee_vat: 0,
        fee_total: 0
      }
    end

    def write_agency_row(sheet, row_num, data)
      sheet.add_cell(row_num, AGENCY_COLUMNS[:media], data[:media])
      sheet.add_cell(row_num, AGENCY_COLUMNS[:advertiser_name], data[:advertiser_name])
      sheet.add_cell(row_num, AGENCY_COLUMNS[:account_id], data[:account_id])
      sheet.add_cell(row_num, AGENCY_COLUMNS[:supply_amount], data[:supply_amount])
      sheet.add_cell(row_num, AGENCY_COLUMNS[:vat], data[:vat])
      sheet.add_cell(row_num, AGENCY_COLUMNS[:total], data[:total])
      sheet.add_cell(row_num, AGENCY_COLUMNS[:fee_rate], data[:fee_rate])
      sheet.add_cell(row_num, AGENCY_COLUMNS[:fee_supply], data[:fee_supply])
      sheet.add_cell(row_num, AGENCY_COLUMNS[:fee_vat], data[:fee_vat])
      sheet.add_cell(row_num, AGENCY_COLUMNS[:fee_total], data[:fee_total])
    end

    def to_number(value)
      return 0 if value.nil?
      return value if value.is_a?(Numeric)
      value.to_s.gsub(/[^0-9.-]/, "").to_f.round
    end
  end
end
