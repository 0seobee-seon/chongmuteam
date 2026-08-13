import sys
import os
sys.path.insert(0, os.path.join(os.path.dirname(__file__), ".."))

import law_api


def test_parse_search_response_multiple_results():
    data = {
        "LawSearch": {
            "law": [
                {
                    "법령일련번호": "001234",
                    "법령명한글": "국민건강보험법",
                    "공포일자": "20250101",
                    "시행일자": "20250701",
                    "소관부처명": "보건복지부",
                    "법령구분명": "법률",
                },
                {
                    "법령일련번호": "005678",
                    "법령명한글": "국민건강보험법 시행령",
                    "공포일자": "20250201",
                    "시행일자": "20250801",
                    "소관부처명": "보건복지부",
                    "법령구분명": "대통령령",
                },
            ]
        }
    }

    result = law_api.parse_search_response(data)

    assert len(result) == 2
    assert result[0] == {
        "mst": "001234",
        "법령명": "국민건강보험법",
        "공포일자": "20250101",
        "시행일자": "20250701",
        "소관부처": "보건복지부",
        "법령구분": "법률",
    }
    assert result[1]["법령명"] == "국민건강보험법 시행령"


def test_parse_search_response_single_result_is_dict_not_list():
    # 검색 결과가 1건일 때 API가 law를 리스트가 아닌 dict로 주는 경우를 처리해야 한다.
    data = {
        "LawSearch": {
            "law": {
                "법령일련번호": "001234",
                "법령명한글": "국민건강보험법",
                "공포일자": "20250101",
                "시행일자": "20250701",
                "소관부처명": "보건복지부",
                "법령구분명": "법률",
            }
        }
    }

    result = law_api.parse_search_response(data)

    assert len(result) == 1
    assert result[0]["법령명"] == "국민건강보험법"


def test_parse_search_response_no_results():
    data = {"LawSearch": {}}

    result = law_api.parse_search_response(data)

    assert result == []


def test_parse_law_response_returns_articles():
    data = {
        "법령": {
            "기본정보": {
                "법령명_한글": "국민건강보험법",
                "공포일자": "20250101",
                "시행일자": "20250701",
            },
            "조문": {
                "조문단위": [
                    {"조문번호": "1", "조문제목": "목적", "조문내용": "제1조(목적) 이 법은 ..."},
                    {"조문번호": "2", "조문제목": "정의", "조문내용": "제2조(정의) 이 법에서..."},
                ]
            },
        }
    }

    result = law_api.parse_law_response(data)

    assert result["법령명"] == "국민건강보험법"
    assert result["공포일자"] == "20250101"
    assert result["시행일자"] == "20250701"
    assert len(result["조문목록"]) == 2
    assert result["조문목록"][0]["조문제목"] == "목적"


def test_parse_law_response_single_article_is_dict_not_list():
    data = {
        "법령": {
            "기본정보": {"법령명_한글": "최저임금법", "공포일자": "20250101", "시행일자": "20250701"},
            "조문": {"조문단위": {"조문번호": "1", "조문제목": "목적", "조문내용": "제1조(목적) ..."}},
        }
    }

    result = law_api.parse_law_response(data)

    assert len(result["조문목록"]) == 1


def test_parse_law_response_article_without_title_key():
    # 장(章) 제목 등 실제 조문이 아닌 항목은 '조문제목' 키 자체가 없을 수 있다.
    data = {
        "법령": {
            "기본정보": {"법령명_한글": "근로기준법", "공포일자": "20250101", "시행일자": "20250701"},
            "조문": {"조문단위": [{"조문번호": "1", "조문내용": "제1장 총칙"}]},
        }
    }

    result = law_api.parse_law_response(data)

    assert result["조문목록"][0]["조문제목"] == ""
    assert result["조문목록"][0]["조문내용"] == "제1장 총칙"


def test_get_law_text_uses_cache_on_second_call(monkeypatch):
    law_api.clear_cache()
    call_count = {"n": 0}

    class FakeResponse:
        def raise_for_status(self):
            pass

        def json(self):
            call_count["n"] += 1
            return {
                "법령": {
                    "기본정보": {"법령명_한글": "근로기준법", "공포일자": "20250101", "시행일자": "20250701"},
                    "조문": {"조문단위": []},
                }
            }

    def fake_get(url, params=None, timeout=None):
        return FakeResponse()

    monkeypatch.setattr(law_api.requests, "get", fake_get)

    law_api.get_law_text("999999", "test-oc")
    law_api.get_law_text("999999", "test-oc")

    assert call_count["n"] == 1  # 두 번째 호출은 캐시를 써서 실제 요청이 안 나가야 한다
