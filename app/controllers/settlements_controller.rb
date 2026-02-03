class SettlementsController < ApplicationController
  before_action :set_settlement, only: [ :show, :process_settlement, :download ]

  def index
    @settlements = Settlement.recent.limit(20)
  end

  def new
    @settlement = Settlement.new
  end

  def create
    @settlement = Settlement.new(settlement_params)

    ActiveRecord::Base.transaction do
      @settlement.save!

      # 템플릿 파일 첨부
      if params[:template_file].present?
        @settlement.template_file.attach(params[:template_file])
      end

      # 정산 파일들 첨부
      if params[:raw_files].present?
        params[:raw_files].each do |file|
          next if file.blank?

          file_type = SettlementFile.detect_file_type(file.original_filename)

          sf = @settlement.settlement_files.create!(
            original_filename: file.original_filename,
            file_type: file_type
          )
          sf.file.attach(file)
        end
      end
    end

    redirect_to @settlement, notice: "정산 작업이 생성되었습니다. '정산 시작' 버튼을 클릭하여 처리를 시작하세요."
  rescue ActiveRecord::RecordInvalid => e
    @settlement ||= Settlement.new
    flash.now[:alert] = "생성 실패: #{e.message}"
    render :new, status: :unprocessable_entity
  end

  def show
    @settlement_files = @settlement.settlement_files.order(:created_at)
  end

  def process_settlement
    if @settlement.pending?
      # 동기 처리 (Sidekiq 없이)
      processor = SettlementProcessor.new(@settlement)
      if processor.process!
        redirect_to @settlement, notice: "정산 처리가 완료되었습니다."
      else
        redirect_to @settlement, alert: "정산 처리 실패: #{processor.errors.join(', ')}"
      end
    else
      redirect_to @settlement, alert: "이미 처리 중이거나 완료된 정산입니다."
    end
  end

  def download
    if @settlement.result_file.attached?
      redirect_to rails_blob_path(@settlement.result_file, disposition: "attachment")
    else
      redirect_to @settlement, alert: "결과 파일이 없습니다."
    end
  end

  private

  def set_settlement
    @settlement = Settlement.find(params[:id])
  end

  def settlement_params
    params.require(:settlement).permit(:period)
  end
end
