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

    attr_reader :workbook, :errors, :distribution_stats

    def initialize(template_path, output_path)
      @template_path = template_path
      @output_path = output_path
      @workbook = nil
      @errors = []
      @distribution_stats = Hash.new { |h, k| h[k] = { rows: 0, accounts: [] } }
      @sheet_row_counters = {}
    end

    def load!
      @workbook = RubyXL::Parser.parse(@template_path)
      initialize_row_counters
      true
    rescue StandardError => e
      @errors << "템플릿 로드 실패: #{e.message}"
      false
    end

    def append_to_agency_sheet(agency_name, row_data, source_type: nil)
      sheet = find_or_create_sheet(agency_name)

      unless sheet
        @errors << "시트를 찾을 수 없음: #{agency_name}"
        return false
      end

      # 현재 행 번호 가져오기
      row_num = next_row_for_sheet(agency_name)

      # 데이터 쓰기
      row_data.each_with_index do |cell_value, col_idx|
        sheet.add_cell(row_num, col_idx, cell_value)
      end

      # 통계 업데이트
      @distribution_stats[agency_name][:rows] += 1
      account_id = extract_account_id_from_row(row_data, source_type)
      @distribution_stats[agency_name][:accounts] << account_id if account_id

      true
    end

    def append_to_media_sheet(media_type, row_data)
      sheet_name = MEDIA_SHEETS[media_type.to_sym]
      return false unless sheet_name

      sheet = find_sheet(sheet_name)
      return false unless sheet

      row_num = next_row_for_sheet(sheet_name)

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

    def find_or_create_sheet(name)
      sheet = find_sheet(name)
      return sheet if sheet

      # 시트가 없으면 새로 생성 (실제 운영시에는 에러 처리가 더 적절할 수 있음)
      # @workbook.add_worksheet(name)
      nil
    end

    def initialize_row_counters
      @workbook.worksheets.each do |sheet|
        name = sheet.sheet_name
        # 기존 데이터의 마지막 행 찾기
        last_row = find_last_row(sheet)
        @sheet_row_counters[name] = last_row + 1
      end
    end

    def find_last_row(sheet)
      return 0 if sheet.sheet_data.nil?

      last_row = 0
      sheet.sheet_data.rows.each_with_index do |row, idx|
        next if row.nil?
        last_row = idx if row.cells&.any? { |c| c&.value.present? }
      end
      last_row
    end

    def next_row_for_sheet(sheet_name)
      @sheet_row_counters[sheet_name] ||= 1
      row = @sheet_row_counters[sheet_name]
      @sheet_row_counters[sheet_name] += 1
      row
    end

    def extract_account_id_from_row(row_data, source_type)
      case source_type.to_s
      when "naver"
        row_data[7]&.to_s # H열
      when /kakao/
        row_data[2]&.to_s # C열
      when "google"
        row_data[1]&.to_s # B열
      else
        nil
      end
    end
  end
end
