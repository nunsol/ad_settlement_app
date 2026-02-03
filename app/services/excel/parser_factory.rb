module Excel
  class ParserFactory
    PARSERS = {
      "naver" => NaverParser,
      "kakao_moment" => -> (path) { KakaoParser.new(path, kakao_type: :moment) },
      "kakao_keyword" => -> (path) { KakaoParser.new(path, kakao_type: :keyword) },
      "kakao_brand" => -> (path) { KakaoParser.new(path, kakao_type: :brand) },
      "kakao_channel" => -> (path) { KakaoParser.new(path, kakao_type: :channel) },
      "google" => GoogleParser
    }.freeze

    def self.create(file_type, file_path)
      parser_class = PARSERS[file_type]

      return nil unless parser_class

      if parser_class.is_a?(Proc)
        parser_class.call(file_path)
      else
        parser_class.new(file_path)
      end
    end

    def self.detect_and_create(filename, file_path)
      file_type = detect_file_type(filename)
      create(file_type, file_path)
    end

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
        nil
      end
    end
  end
end
