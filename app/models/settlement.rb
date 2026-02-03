class Settlement < ApplicationRecord
  has_many :settlement_files, dependent: :destroy
  has_one_attached :template_file
  has_one_attached :result_file

  STATUSES = %w[pending processing completed failed].freeze

  validates :period, presence: true
  validates :status, inclusion: { in: STATUSES }

  before_validation :set_default_status, on: :create

  scope :recent, -> { order(created_at: :desc) }

  def pending?
    status == "pending"
  end

  def processing?
    status == "processing"
  end

  def completed?
    status == "completed"
  end

  def failed?
    status == "failed"
  end

  def can_process?
    pending? && template_file.attached? && settlement_files.any?
  end

  def match_rate
    return 0 if total_rows.to_i.zero?
    (matched_rows.to_f / total_rows * 100).round(1)
  end

  private

  def set_default_status
    self.status ||= "pending"
  end
end
