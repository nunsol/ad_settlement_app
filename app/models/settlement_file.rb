class SettlementFile < ApplicationRecord
  belongs_to :settlement
  has_one_attached :file

  FILE_TYPES = %w[naver kakao_moment kakao_keyword kakao_brand kakao_channel google unknown].freeze
  STATUSES = %w[pending processing completed failed].freeze

  validates :file_type, inclusion: { in: FILE_TYPES }
  validates :status, inclusion: { in: STATUSES }
  validates :original_filename, presence: true

  before_validation :set_defaults, on: :create

  def self.detect_file_type(filename)
    case filename
    when /네이버SA_GFA|네이버_SA|naver/i
      "naver"
    when /카카오_모먼트|카카오모먼트|kakao.*moment/i
      "kakao_moment"
    when /카카오_키워드|카카오키워드|kakao.*keyword/i
      "kakao_keyword"
    when /카카오_브랜드|카카오브랜드|kakao.*brand/i
      "kakao_brand"
    when /카카오톡.*채널|카카오채널|kakao.*channel/i
      "kakao_channel"
    when /구글정산|구글.*정산|google/i
      "google"
    else
      "unknown"
    end
  end

  private

  def set_defaults
    self.status ||= "pending"
    self.file_type ||= "unknown"
  end
end
