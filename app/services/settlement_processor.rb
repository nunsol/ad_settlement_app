class SettlementProcessor
  attr_reader :settlement, :errors, :stats

  def initialize(settlement)
    @settlement = settlement
    @errors = []
    @stats = {
      total_rows: 0,
      matched_rows: 0,
      unmatched_rows: 0,
      unmatched_accounts: [],
      agency_distribution: {}
    }
  end

  def process!
    return false unless validate_settlement

    @settlement.update!(status: "processing")

    begin
      # 1. 템플릿 파일 준비
      template_path = download_to_temp(@settlement.template_file)
      output_path = generate_output_path

      # 2. 계정 라우터 초기화
      router = Routing::AccountRouter.new(template_path)
      unless router.load!
        raise ProcessingError, "계정정리 시트 로드 실패: #{router.errors.join(', ')}"
      end

      Rails.logger.info "AccountRouter loaded: #{router.total_accounts} accounts"

      # 3. 템플릿 프로세서 초기화
      processor = Excel::TemplateProcessor.new(template_path, output_path)
      unless processor.load!
        raise ProcessingError, "템플릿 로드 실패: #{processor.errors.join(', ')}"
      end

      # 4. 각 정산 파일 처리
      @settlement.settlement_files.each do |sf|
        process_settlement_file(sf, router, processor)
      end

      # 5. 결과 저장
      processor.save!

      # 6. 결과 파일 첨부
      @settlement.result_file.attach(
        io: File.open(output_path),
        filename: "정산결과_#{@settlement.period}_#{Time.current.strftime('%Y%m%d%H%M%S')}.xlsx",
        content_type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
      )

      # 7. 통계 업데이트
      update_settlement_stats(processor)

      @settlement.update!(status: "completed")
      true

    rescue StandardError => e
      Rails.logger.error "Settlement processing failed: #{e.message}\n#{e.backtrace.join("\n")}"
      @errors << e.message
      @settlement.update!(status: "failed", error_message: e.message)
      false
    ensure
      cleanup_temp_files
    end
  end

  private

  def validate_settlement
    unless @settlement.template_file.attached?
      @errors << "템플릿 파일이 없습니다"
      return false
    end

    unless @settlement.settlement_files.any?
      @errors << "정산 파일이 없습니다"
      return false
    end

    true
  end

  def download_to_temp(attachment)
    temp_path = Rails.root.join("tmp", "settlement_#{@settlement.id}_template.xlsx")
    File.open(temp_path, "wb") do |file|
      file.write(attachment.download)
    end
    @temp_files ||= []
    @temp_files << temp_path
    temp_path.to_s
  end

  def generate_output_path
    path = Rails.root.join("tmp", "settlement_#{@settlement.id}_result.xlsx")
    @temp_files ||= []
    @temp_files << path
    path.to_s
  end

  def process_settlement_file(settlement_file, router, processor)
    settlement_file.update!(status: "processing")

    begin
      # 파일 다운로드
      file_path = download_settlement_file(settlement_file)

      # 파서 생성
      parser = Excel::ParserFactory.create(settlement_file.file_type, file_path)
      unless parser
        raise ProcessingError, "지원하지 않는 파일 유형: #{settlement_file.file_type}"
      end

      # 데이터 파싱
      result = parser.parse
      if result[:errors].any?
        Rails.logger.warn "Parser errors for #{settlement_file.original_filename}: #{result[:errors]}"
      end

      matched_count = 0
      rows = result[:rows] || []

      rows.each do |row|
        account_id = row[:account_id]
        route_info = router.route(account_id)

        if route_info
          # 대행사 시트에 데이터 추가
          agency = route_info[:agency]
          processor.append_to_agency_sheet(
            agency,
            row[:raw_data],
            source_type: settlement_file.file_type
          )

          # 매체 통합 시트에도 추가
          media_type = settlement_file.file_type.to_sym
          processor.append_to_media_sheet(media_type, row[:raw_data])

          matched_count += 1
          @stats[:matched_rows] += 1
        else
          @stats[:unmatched_rows] += 1
          @stats[:unmatched_accounts] << {
            account_id: account_id,
            file: settlement_file.original_filename,
            row_number: row[:row_number]
          }
        end

        @stats[:total_rows] += 1
      end

      settlement_file.update!(
        status: "completed",
        rows_count: rows.size,
        matched_count: matched_count
      )

    rescue StandardError => e
      settlement_file.update!(status: "failed")
      Rails.logger.error "Failed to process file #{settlement_file.original_filename}: #{e.message}"
      raise
    end
  end

  def download_settlement_file(settlement_file)
    temp_path = Rails.root.join("tmp", "settlement_#{@settlement.id}_file_#{settlement_file.id}.xlsx")
    File.open(temp_path, "wb") do |file|
      file.write(settlement_file.file.download)
    end
    @temp_files ||= []
    @temp_files << temp_path
    temp_path.to_s
  end

  def update_settlement_stats(processor)
    @stats[:agency_distribution] = processor.distribution_stats.transform_values do |v|
      { rows: v[:rows], unique_accounts: v[:accounts].uniq.size }
    end

    @settlement.update!(
      total_rows: @stats[:total_rows],
      matched_rows: @stats[:matched_rows],
      unmatched_rows: @stats[:unmatched_rows],
      unmatched_accounts: @stats[:unmatched_accounts].first(100), # 최대 100개만 저장
      agency_distribution: @stats[:agency_distribution]
    )
  end

  def cleanup_temp_files
    return unless @temp_files

    @temp_files.each do |path|
      File.delete(path) if File.exist?(path)
    end
  end

  class ProcessingError < StandardError; end
end
