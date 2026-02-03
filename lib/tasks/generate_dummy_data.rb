#!/usr/bin/env ruby
# 테스트용 더미 데이터 생성 스크립트

require 'rubyXL'
require 'rubyXL/convenience_methods'

OUTPUT_DIR = File.expand_path('../../dummy', __dir__)

# 대행사 목록 (실제 시트로 생성될 이름들)
AGENCIES = [
  "준커뮤니케이션",
  "스튜디오킥",
  "넥스브이",
  "애드스톤",
  "그로스미디어",
  "토리커뮤니케이션",
  "캐치인마케팅",
  "드림인사이트",
  "메이크메이커",
  "디브컴퍼니"
]

# 계정 데이터 생성 (계정ID, 광고주명, 대행사)
def generate_accounts(count = 100)
  accounts = []
  count.times do |i|
    accounts << {
      account_id: (1000000 + i).to_s,
      advertiser_name: "광고주#{i + 1}",
      account_name: "계정#{i + 1}",
      manager: "담당자#{(i % 10) + 1}",
      media: ["네이버", "카카오", "구글"][i % 3],
      product: ["검색광고", "디스플레이", "브랜드검색"][i % 3],
      agency: AGENCIES[i % AGENCIES.length],
      commission_rate: [0.10, 0.12, 0.15, 0.18, 0.20][i % 5]
    }
  end
  accounts
end

# 마스터 템플릿 생성
def create_master_template(accounts)
  workbook = RubyXL::Workbook.new

  # 첫 번째 시트를 계정정리로 변경
  account_sheet = workbook.worksheets[0]
  account_sheet.sheet_name = "계정정리"

  # 헤더
  headers = ["광고주명", "계정명", "계정아이디(구글CID)", "계정비번", "계산서 발행구분", "", "담당자명", "매체", "상품", "대행사", "수수료율", "차감수수료율", "비고"]
  headers.each_with_index do |h, col|
    account_sheet.add_cell(0, col, h)
  end

  # 계정 데이터
  accounts.each_with_index do |acc, row|
    account_sheet.add_cell(row + 1, 0, acc[:advertiser_name])
    account_sheet.add_cell(row + 1, 1, acc[:account_name])
    account_sheet.add_cell(row + 1, 2, acc[:account_id])
    account_sheet.add_cell(row + 1, 3, "****")
    account_sheet.add_cell(row + 1, 4, "위임")
    account_sheet.add_cell(row + 1, 5, "계정직접충전")
    account_sheet.add_cell(row + 1, 6, acc[:manager])
    account_sheet.add_cell(row + 1, 7, acc[:media])
    account_sheet.add_cell(row + 1, 8, acc[:product])
    account_sheet.add_cell(row + 1, 9, acc[:agency])
    account_sheet.add_cell(row + 1, 10, acc[:commission_rate])
    account_sheet.add_cell(row + 1, 11, acc[:commission_rate] - 0.005)
  end

  # 대행사 시트 생성
  AGENCIES.each do |agency|
    sheet = workbook.add_worksheet(agency)
    # 헤더만 추가 (데이터는 정산 시 채워짐)
    sheet.add_cell(0, 0, "광고주명")
    sheet.add_cell(0, 1, "계정ID")
    sheet.add_cell(0, 2, "매체")
    sheet.add_cell(0, 3, "금액")
  end

  # 매체 통합 시트
  ["네이버통합_SA,GFA", "카카오모먼트", "다음_검색(신)", "다음_브랜드검색", "다음_카카오톡채널", "구글"].each do |name|
    sheet = workbook.add_worksheet(name)
    sheet.add_cell(0, 0, "데이터")
  end

  workbook.write(File.join(OUTPUT_DIR, "마스터템플릿_테스트.xlsx"))
  puts "✓ 마스터템플릿_테스트.xlsx 생성 완료"
end

# 네이버 정산 데이터 생성
def create_naver_data(accounts)
  workbook = RubyXL::Workbook.new
  sheet = workbook.worksheets[0]
  sheet.sheet_name = "25년 10월(SA,GFA실적)"

  # 헤더 (H열이 계정ID)
  headers = ["대행사", "에이전트ID", "담당 에이전트 이름", "소속 관리 계정 ID", "소속 관리 계정 이름", "광고주ID", "광고계정 CustomerID", "광고계정 AdAccountNo", "광고주명", "결제유형", "웹방문(내부/PC)", "CPC 유상", "유상실적TOTAL"]
  headers.each_with_index { |h, col| sheet.add_cell(0, col, h) }

  # 네이버 계정만 필터링해서 데이터 생성
  naver_accounts = accounts.select { |a| a[:media] == "네이버" }

  row = 1
  naver_accounts.each do |acc|
    # 각 계정당 3개 행 생성 (여러 날짜 데이터)
    3.times do |day|
      sheet.add_cell(row, 0, acc[:agency])
      sheet.add_cell(row, 1, "AGT#{rand(1000..9999)}")
      sheet.add_cell(row, 2, acc[:manager])
      sheet.add_cell(row, 3, "MGR#{rand(100..999)}")
      sheet.add_cell(row, 4, "관리계정#{rand(1..10)}")
      sheet.add_cell(row, 5, "ADV#{rand(10000..99999)}")
      sheet.add_cell(row, 6, "CUS#{rand(10000..99999)}")
      sheet.add_cell(row, 7, acc[:account_id])  # H열: 계정 ID (매칭 키)
      sheet.add_cell(row, 8, acc[:advertiser_name])
      sheet.add_cell(row, 9, "후불")
      sheet.add_cell(row, 10, rand(10000..100000))
      sheet.add_cell(row, 11, rand(50000..500000))
      sheet.add_cell(row, 12, rand(100000..1000000))
      row += 1
    end
  end

  workbook.write(File.join(OUTPUT_DIR, "네이버SA_GFA_정산_테스트.xlsx"))
  puts "✓ 네이버SA_GFA_정산_테스트.xlsx 생성 완료 (#{row - 1}행)"
end

# 카카오 모먼트 정산 데이터 생성
def create_kakao_moment_data(accounts)
  workbook = RubyXL::Workbook.new
  sheet = workbook.worksheets[0]
  sheet.sheet_name = "카카오모먼트_202510"

  # 헤더 (C열이 자산ID = 계정ID)
  headers = ["매출일", "서비스", "자산 ID", "자산 이름", "광고주 사업자명", "마케터명", "월렛 ID", "월렛 이름", "유형", "사업자번호", "사업자명", "공급가액", "부가세", "무상캐시"]
  headers.each_with_index { |h, col| sheet.add_cell(0, col, h) }

  # 카카오 계정만 필터링
  kakao_accounts = accounts.select { |a| a[:media] == "카카오" }

  row = 1
  kakao_accounts.each do |acc|
    5.times do |day|
      sheet.add_cell(row, 0, "2025-10-#{(day + 1).to_s.rjust(2, '0')}")
      sheet.add_cell(row, 1, "카카오모먼트")
      sheet.add_cell(row, 2, acc[:account_id])  # C열: 자산 ID (매칭 키)
      sheet.add_cell(row, 3, "자산_#{acc[:account_name]}")
      sheet.add_cell(row, 4, acc[:advertiser_name])
      sheet.add_cell(row, 5, acc[:manager])
      sheet.add_cell(row, 6, acc[:account_id])
      sheet.add_cell(row, 7, "월렛_#{acc[:account_name]}")
      sheet.add_cell(row, 8, "일반")
      sheet.add_cell(row, 9, "123-45-#{rand(10000..99999)}")
      sheet.add_cell(row, 10, acc[:advertiser_name])
      sheet.add_cell(row, 11, rand(50000..300000))
      sheet.add_cell(row, 12, rand(5000..30000))
      sheet.add_cell(row, 13, rand(0..10000))
      row += 1
    end
  end

  workbook.write(File.join(OUTPUT_DIR, "카카오_모먼트_정산_테스트.xlsx"))
  puts "✓ 카카오_모먼트_정산_테스트.xlsx 생성 완료 (#{row - 1}행)"
end

# 카카오 키워드 정산 데이터 생성
def create_kakao_keyword_data(accounts)
  workbook = RubyXL::Workbook.new
  sheet = workbook.worksheets[0]
  sheet.sheet_name = "카카오 키워드_202510"

  headers = ["매출일", "서비스", "자산 ID", "자산 이름", "광고주 사업자명", "마케터명", "월렛 ID", "월렛 이름", "유형", "사업자번호", "사업자명", "공급가액", "부가세", "무상캐시"]
  headers.each_with_index { |h, col| sheet.add_cell(0, col, h) }

  kakao_accounts = accounts.select { |a| a[:media] == "카카오" }

  row = 1
  kakao_accounts.first(15).each do |acc|
    3.times do |day|
      sheet.add_cell(row, 0, "2025-10-#{(day + 1).to_s.rjust(2, '0')}")
      sheet.add_cell(row, 1, "키워드광고")
      sheet.add_cell(row, 2, acc[:account_id])
      sheet.add_cell(row, 3, "자산_#{acc[:account_name]}")
      sheet.add_cell(row, 4, acc[:advertiser_name])
      sheet.add_cell(row, 5, acc[:manager])
      sheet.add_cell(row, 6, acc[:account_id])
      sheet.add_cell(row, 7, "월렛_#{acc[:account_name]}")
      sheet.add_cell(row, 8, "일반")
      sheet.add_cell(row, 9, "123-45-#{rand(10000..99999)}")
      sheet.add_cell(row, 10, acc[:advertiser_name])
      sheet.add_cell(row, 11, rand(30000..200000))
      sheet.add_cell(row, 12, rand(3000..20000))
      sheet.add_cell(row, 13, rand(0..5000))
      row += 1
    end
  end

  workbook.write(File.join(OUTPUT_DIR, "카카오_키워드정산_테스트.xlsx"))
  puts "✓ 카카오_키워드정산_테스트.xlsx 생성 완료 (#{row - 1}행)"
end

# 구글 정산 데이터 생성
def create_google_data(accounts)
  workbook = RubyXL::Workbook.new
  sheet = workbook.worksheets[0]
  sheet.sheet_name = "25.10"

  # 구글은 특수한 형식
  sheet.add_cell(1, 1, "대대행사 수수료 금액")
  sheet.add_cell(2, 1, "구분")
  sheet.add_cell(2, 2, "매체")
  sheet.add_cell(2, 3, "수수료율")
  sheet.add_cell(2, 4, "수수료 공급가액")

  # 헤더 행 (11행, 인덱스 10)
  sheet.add_cell(11, 0, "구글\n10월")
  sheet.add_cell(11, 1, "계정 ID")
  sheet.add_cell(11, 2, "담당자")
  sheet.add_cell(11, 3, "계정")
  sheet.add_cell(11, 4, "공급가액")
  sheet.add_cell(11, 5, "DA 공급가액")

  # 구글 계정만 필터링
  google_accounts = accounts.select { |a| a[:media] == "구글" }

  row = 12
  google_accounts.each_with_index do |acc, idx|
    sheet.add_cell(row, 0, nil)
    sheet.add_cell(row, 1, idx + 1)  # 순번
    sheet.add_cell(row, 2, "A-후불")
    sheet.add_cell(row, 3, acc[:account_id])  # D열: 계정 ID (매칭 키)
    sheet.add_cell(row, 4, rand(1000000..10000000))
    sheet.add_cell(row, 5, rand(100000..1000000))
    row += 1
  end

  workbook.write(File.join(OUTPUT_DIR, "구글정산(수수료포함)_테스트.xlsx"))
  puts "✓ 구글정산(수수료포함)_테스트.xlsx 생성 완료 (#{row - 12}행)"
end

# 메인 실행
puts "=" * 50
puts "테스트용 더미 데이터 생성 시작"
puts "출력 폴더: #{OUTPUT_DIR}"
puts "=" * 50

accounts = generate_accounts(100)
puts "\n총 #{accounts.length}개 계정 생성"
puts "대행사별 분포:"
AGENCIES.each do |agency|
  count = accounts.count { |a| a[:agency] == agency }
  puts "  - #{agency}: #{count}개"
end

puts "\n파일 생성 중..."
create_master_template(accounts)
create_naver_data(accounts)
create_kakao_moment_data(accounts)
create_kakao_keyword_data(accounts)
create_google_data(accounts)

puts "\n" + "=" * 50
puts "더미 데이터 생성 완료!"
puts "=" * 50
