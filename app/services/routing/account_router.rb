module Routing
  class AccountRouter
    ACCOUNT_SHEET_NAME = "계정정리"

    # 계정정리 시트 컬럼 인덱스 (0-based)
    COLUMNS = {
      advertiser_name: 0,    # A열: 광고주명
      account_name: 1,       # B열: 계정명
      account_id: 2,         # C열: 계정아이디(구글CID)
      account_password: 3,   # D열: 계정비번
      invoice_type: 4,       # E열: 계산서 발행구분
      charge_type: 5,        # F열: (충전유형)
      manager: 6,            # G열: 담당자명
      media: 7,              # H열: 매체
      product: 8,            # I열: 상품
      agency: 9,             # J열: 대행사
      commission_rate: 10,   # K열: 수수료율
      deduction_rate: 11,    # L열: 차감수수료율
      note: 12               # M열: 비고
    }.freeze

    attr_reader :account_map, :agency_list, :errors

    def initialize(template_path)
      @template_path = template_path
      @account_map = {}
      @agency_list = []
      @errors = []
    end

    def load!
      workbook = Roo::Spreadsheet.open(@template_path)

      unless workbook.sheets.include?(ACCOUNT_SHEET_NAME)
        @errors << "계정정리 시트를 찾을 수 없습니다"
        return false
      end

      sheet = workbook.sheet(ACCOUNT_SHEET_NAME)
      @agency_list = detect_agency_sheets(workbook)

      # 계정정리 시트 파싱
      (2..sheet.last_row).each do |row_num|
        row = sheet.row(row_num)
        next if row.compact.empty?

        account_id = normalize_account_id(row[COLUMNS[:account_id]])
        next if account_id.nil? || account_id.empty?

        agency = row[COLUMNS[:agency]]&.to_s&.strip
        next if agency.nil? || agency.empty?

        @account_map[account_id] = {
          advertiser_name: row[COLUMNS[:advertiser_name]]&.to_s&.strip,
          account_name: row[COLUMNS[:account_name]]&.to_s&.strip,
          account_id: account_id,
          manager: row[COLUMNS[:manager]]&.to_s&.strip,
          media: row[COLUMNS[:media]]&.to_s&.strip,
          product: row[COLUMNS[:product]]&.to_s&.strip,
          agency: agency,
          commission_rate: row[COLUMNS[:commission_rate]],
          deduction_rate: row[COLUMNS[:deduction_rate]]
        }
      end

      true
    rescue StandardError => e
      @errors << "템플릿 파일 로드 실패: #{e.message}"
      false
    end

    def route(account_id)
      normalized_id = normalize_account_id(account_id)
      account_info = @account_map[normalized_id]

      return nil unless account_info

      agency = account_info[:agency]

      # 대행사 시트가 존재하는지 확인
      if @agency_list.include?(agency)
        { agency: agency, account_info: account_info }
      else
        # 유사한 대행사명 검색
        similar = find_similar_agency(agency)
        if similar
          { agency: similar, account_info: account_info, original_agency: agency }
        else
          nil
        end
      end
    end

    def find_agency(account_id)
      route_info = route(account_id)
      route_info&.dig(:agency)
    end

    def loaded?
      @account_map.any?
    end

    def total_accounts
      @account_map.size
    end

    def agency_stats
      stats = Hash.new(0)
      @account_map.each_value do |info|
        stats[info[:agency]] += 1
      end
      stats.sort_by { |_, count| -count }.to_h
    end

    private

    def normalize_account_id(value)
      return nil if value.nil?
      value.to_s.strip.gsub(/[^0-9a-zA-Z\-_]/, "")
    end

    def detect_agency_sheets(workbook)
      # 특수 시트 제외하고 대행사 시트 목록 추출
      special_sheets = [
        "계정정리", "네이버통합_SA,GFA", "다음_검색(신)", "다음_검색",
        "카카오모먼트", "다음_브랜드검색", "다음_카카오톡채널",
        "구글", "구글_디플랜360", "세금계산서 발행 내역", "Sheet1"
      ]

      workbook.sheets.reject { |s| special_sheets.include?(s) }
    end

    def find_similar_agency(agency_name)
      return nil if agency_name.nil?

      # 정확한 매칭 시도
      return agency_name if @agency_list.include?(agency_name)

      # 부분 매칭 시도
      @agency_list.find do |sheet_name|
        sheet_name.include?(agency_name) || agency_name.include?(sheet_name)
      end
    end
  end
end
