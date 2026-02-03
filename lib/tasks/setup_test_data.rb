#!/usr/bin/env ruby
# 마스터 템플릿 기반 테스트 데이터 생성 스크립트
# 사용자의 정산마스터템플릿_더미데이터.xlsx를 기반으로 테스트 데이터를 생성합니다.

require 'roo'
require 'rubyXL'
require 'rubyXL/convenience_methods'
require 'fileutils'

PROJECT_ROOT = File.expand_path('../..', __dir__)
MASTER_TEMPLATE = File.join(PROJECT_ROOT, '정산마스터템플릿_더미데이터.xlsx')
OUTPUT_DIR = File.join(PROJECT_ROOT, 'dummy')

FileUtils.mkdir_p(OUTPUT_DIR)

puts "=" * 60
puts "마스터 템플릿 기반 테스트 데이터 생성"
puts "=" * 60

# 마스터 템플릿 읽기
unless File.exist?(MASTER_TEMPLATE)
  puts "오류: 마스터 템플릿 파일을 찾을 수 없습니다."
  puts "경로: #{MASTER_TEMPLATE}"
  exit 1
end

puts "\n1. 마스터 템플릿 분석 중..."
xlsx = Roo::Spreadsheet.open(MASTER_TEMPLATE)

# 대행사 시트 목록 추출 (특수 시트 제외)
special_sheets = [
  "계정정리", "네이버통합_SA,GFA", "다음_검색(신)", "다음_검색",
  "카카오모먼트", "다음_브랜드검색", "다음_카카오톡채널",
  "구글", "구글_디플랜360", "세금계산서 발행 내역", "Sheet1"
]
agency_sheets = xlsx.sheets.reject { |s| special_sheets.include?(s) }

puts "   - 총 시트 수: #{xlsx.sheets.count}"
puts "   - 대행사 시트 수: #{agency_sheets.count}"

# 테스트에 사용할 대행사 선택 (10개)
test_agencies = agency_sheets.first(10)
puts "   - 테스트용 대행사: #{test_agencies.join(', ')}"

# 계정 데이터 생성 (각 대행사당 10개 계정)
puts "\n2. 계정 데이터 생성 중..."
accounts = []
account_id_base = 1000000

test_agencies.each_with_index do |agency, agency_idx|
  10.times do |i|
    account_id = (account_id_base + agency_idx * 10 + i).to_s
    accounts << {
      account_id: account_id,
      advertiser_name: "광고주#{accounts.length + 1}",
      account_name: "계정#{accounts.length + 1}",
      manager: "담당자#{(i % 5) + 1}",
      media: ["네이버", "네이버", "네이버", "카카오", "구글"][i % 5],  # 네이버 비중 높게
      product: ["검색광고", "GFA", "브랜드검색", "모먼트", "검색"][i % 5],
      agency: agency,
      commission_rate: [0.10, 0.12, 0.15, 0.18, 0.20][i % 5]
    }
  end
end

puts "   - 총 계정 수: #{accounts.length}"

# 매체별 분포
media_dist = accounts.group_by { |a| a[:media] }
puts "   - 매체별 분포:"
media_dist.each { |m, arr| puts "     #{m}: #{arr.length}개" }

# 마스터 템플릿 복사 및 계정정리 시트 업데이트
puts "\n3. 마스터 템플릿 복사 및 업데이트 중..."
output_template = File.join(OUTPUT_DIR, '마스터템플릿_테스트.xlsx')

# RubyXL로 템플릿 열기 (수정 가능)
workbook = RubyXL::Parser.parse(MASTER_TEMPLATE)

# 계정정리 시트 찾기
account_sheet = workbook['계정정리']
unless account_sheet
  puts "오류: 계정정리 시트를 찾을 수 없습니다."
  exit 1
end

# 기존 데이터 삭제 (2행부터)
# 계정정리 시트 컬럼: A(빈칸) B(광고주명) C(계정명) D(계정ID) E(계정비번) F(계산서발행구분) G(충전유형) H(담당자명) I(매체) J(상품) K(대행사) L(수수료율) M(차감수수료율) N(비고)
(2..account_sheet.sheet_data.rows.size).each do |row_idx|
  (1..14).each do |col_idx|
    cell = account_sheet[row_idx - 1]&.cells&.[](col_idx)
    cell&.change_contents(nil)
  end
end

# 새 계정 데이터 입력 (헤더가 1행, 데이터는 2행부터)
accounts.each_with_index do |acc, idx|
  row_idx = idx + 1  # 0-indexed, 2행부터
  account_sheet.add_cell(row_idx, 1, acc[:advertiser_name])   # B열
  account_sheet.add_cell(row_idx, 2, acc[:account_name])      # C열
  account_sheet.add_cell(row_idx, 3, acc[:account_id])        # D열
  account_sheet.add_cell(row_idx, 4, "****")                  # E열
  account_sheet.add_cell(row_idx, 5, "위임")                  # F열
  account_sheet.add_cell(row_idx, 6, "계정직접충전")          # G열
  account_sheet.add_cell(row_idx, 7, acc[:manager])           # H열
  account_sheet.add_cell(row_idx, 8, acc[:media])             # I열
  account_sheet.add_cell(row_idx, 9, acc[:product])           # J열
  account_sheet.add_cell(row_idx, 10, acc[:agency])           # K열
  account_sheet.add_cell(row_idx, 11, acc[:commission_rate])  # L열
  account_sheet.add_cell(row_idx, 12, acc[:commission_rate] - 0.005)  # M열
end

workbook.write(output_template)
puts "   - 저장: #{output_template}"

# 네이버 Raw 데이터 생성
puts "\n4. 네이버 SA/GFA 정산 데이터 생성 중..."
naver_workbook = RubyXL::Workbook.new
naver_sheet = naver_workbook.worksheets[0]
naver_sheet.sheet_name = "25년 10월(SA,GFA실적)"

# 헤더
naver_headers = ["대행사", "에이전트ID", "담당 에이전트 이름", "소속 관리 계정 ID", "소속 관리 계정 이름", "광고주ID", "광고계정 CustomerID", "광고계정 AdAccountNo", "광고주명", "결제유형", "웹방문(내부/PC)", "CPC 유상", "유상실적TOTAL"]
naver_headers.each_with_index { |h, col| naver_sheet.add_cell(0, col, h) }

naver_accounts = accounts.select { |a| a[:media] == "네이버" }
row = 1
naver_accounts.each do |acc|
  3.times do |day|
    naver_sheet.add_cell(row, 0, acc[:agency])
    naver_sheet.add_cell(row, 1, "AGT#{rand(1000..9999)}")
    naver_sheet.add_cell(row, 2, acc[:manager])
    naver_sheet.add_cell(row, 3, "MGR#{rand(100..999)}")
    naver_sheet.add_cell(row, 4, "관리계정#{rand(1..10)}")
    naver_sheet.add_cell(row, 5, "ADV#{rand(10000..99999)}")
    naver_sheet.add_cell(row, 6, "CUS#{rand(10000..99999)}")
    naver_sheet.add_cell(row, 7, acc[:account_id])  # H열: 계정 ID
    naver_sheet.add_cell(row, 8, acc[:advertiser_name])
    naver_sheet.add_cell(row, 9, "후불")
    naver_sheet.add_cell(row, 10, rand(10000..100000))
    naver_sheet.add_cell(row, 11, rand(50000..500000))
    naver_sheet.add_cell(row, 12, rand(100000..1000000))
    row += 1
  end
end

naver_file = File.join(OUTPUT_DIR, '네이버SA_GFA_정산_테스트.xlsx')
naver_workbook.write(naver_file)
puts "   - 저장: #{naver_file} (#{row - 1}행)"

# 카카오 모먼트 Raw 데이터 생성
puts "\n5. 카카오 모먼트 정산 데이터 생성 중..."
kakao_workbook = RubyXL::Workbook.new
kakao_sheet = kakao_workbook.worksheets[0]
kakao_sheet.sheet_name = "카카오모먼트_202510"

kakao_headers = ["매출일", "서비스", "자산 ID", "자산 이름", "광고주 사업자명", "마케터명", "월렛 ID", "월렛 이름", "유형", "사업자번호", "사업자명", "공급가액", "부가세", "무상캐시"]
kakao_headers.each_with_index { |h, col| kakao_sheet.add_cell(0, col, h) }

kakao_accounts = accounts.select { |a| a[:media] == "카카오" }
row = 1
kakao_accounts.each do |acc|
  5.times do |day|
    kakao_sheet.add_cell(row, 0, "2025-10-#{(day + 1).to_s.rjust(2, '0')}")
    kakao_sheet.add_cell(row, 1, "카카오모먼트")
    kakao_sheet.add_cell(row, 2, acc[:account_id])  # C열: 자산 ID
    kakao_sheet.add_cell(row, 3, "자산_#{acc[:account_name]}")
    kakao_sheet.add_cell(row, 4, acc[:advertiser_name])
    kakao_sheet.add_cell(row, 5, acc[:manager])
    kakao_sheet.add_cell(row, 6, acc[:account_id])
    kakao_sheet.add_cell(row, 7, "월렛_#{acc[:account_name]}")
    kakao_sheet.add_cell(row, 8, "일반")
    kakao_sheet.add_cell(row, 9, "123-45-#{rand(10000..99999)}")
    kakao_sheet.add_cell(row, 10, acc[:advertiser_name])
    kakao_sheet.add_cell(row, 11, rand(50000..300000))
    kakao_sheet.add_cell(row, 12, rand(5000..30000))
    kakao_sheet.add_cell(row, 13, rand(0..10000))
    row += 1
  end
end

kakao_file = File.join(OUTPUT_DIR, '카카오_모먼트_정산_테스트.xlsx')
kakao_workbook.write(kakao_file)
puts "   - 저장: #{kakao_file} (#{row - 1}행)"

# 카카오 키워드 Raw 데이터 생성
puts "\n6. 카카오 키워드 정산 데이터 생성 중..."
keyword_workbook = RubyXL::Workbook.new
keyword_sheet = keyword_workbook.worksheets[0]
keyword_sheet.sheet_name = "카카오 키워드_202510"

kakao_headers.each_with_index { |h, col| keyword_sheet.add_cell(0, col, h) }

row = 1
kakao_accounts.first(10).each do |acc|
  3.times do |day|
    keyword_sheet.add_cell(row, 0, "2025-10-#{(day + 1).to_s.rjust(2, '0')}")
    keyword_sheet.add_cell(row, 1, "키워드광고")
    keyword_sheet.add_cell(row, 2, acc[:account_id])
    keyword_sheet.add_cell(row, 3, "자산_#{acc[:account_name]}")
    keyword_sheet.add_cell(row, 4, acc[:advertiser_name])
    keyword_sheet.add_cell(row, 5, acc[:manager])
    keyword_sheet.add_cell(row, 6, acc[:account_id])
    keyword_sheet.add_cell(row, 7, "월렛_#{acc[:account_name]}")
    keyword_sheet.add_cell(row, 8, "일반")
    keyword_sheet.add_cell(row, 9, "123-45-#{rand(10000..99999)}")
    keyword_sheet.add_cell(row, 10, acc[:advertiser_name])
    keyword_sheet.add_cell(row, 11, rand(30000..200000))
    keyword_sheet.add_cell(row, 12, rand(3000..20000))
    keyword_sheet.add_cell(row, 13, rand(0..5000))
    row += 1
  end
end

keyword_file = File.join(OUTPUT_DIR, '카카오_키워드정산_테스트.xlsx')
keyword_workbook.write(keyword_file)
puts "   - 저장: #{keyword_file} (#{row - 1}행)"

# 구글 Raw 데이터 생성
puts "\n7. 구글 정산 데이터 생성 중..."
google_workbook = RubyXL::Workbook.new
google_sheet = google_workbook.worksheets[0]
google_sheet.sheet_name = "25.10"

# 구글 특수 형식
google_sheet.add_cell(1, 1, "대대행사 수수료 금액")
google_sheet.add_cell(2, 1, "구분")
google_sheet.add_cell(2, 2, "매체")
google_sheet.add_cell(2, 3, "수수료율")
google_sheet.add_cell(2, 4, "수수료 공급가액")

# 헤더 행 (12행, 인덱스 11)
google_sheet.add_cell(11, 0, "구글\n10월")
google_sheet.add_cell(11, 1, "계정 ID")
google_sheet.add_cell(11, 2, "담당자")
google_sheet.add_cell(11, 3, "계정")
google_sheet.add_cell(11, 4, "공급가액")
google_sheet.add_cell(11, 5, "DA 공급가액")

google_accounts = accounts.select { |a| a[:media] == "구글" }
row = 12
google_accounts.each_with_index do |acc, idx|
  google_sheet.add_cell(row, 0, nil)
  google_sheet.add_cell(row, 1, idx + 1)
  google_sheet.add_cell(row, 2, "A-후불")
  google_sheet.add_cell(row, 3, acc[:account_id])  # D열: 계정 ID
  google_sheet.add_cell(row, 4, rand(1000000..10000000))
  google_sheet.add_cell(row, 5, rand(100000..1000000))
  row += 1
end

google_file = File.join(OUTPUT_DIR, '구글정산(수수료포함)_테스트.xlsx')
google_workbook.write(google_file)
puts "   - 저장: #{google_file} (#{row - 12}행)"

puts "\n" + "=" * 60
puts "테스트 데이터 생성 완료!"
puts "=" * 60
puts "\n생성된 파일:"
puts "  - #{output_template}"
puts "  - #{naver_file}"
puts "  - #{kakao_file}"
puts "  - #{keyword_file}"
puts "  - #{google_file}"
