"""
국가법령정보 공동활용 API 클라이언트
공식 문서: https://open.law.go.kr
"""

import requests

BASE_URL = "http://www.law.go.kr/DRF"
TIMEOUT = 10

_text_cache = {}  # mst -> 본문 dict. 세션(프로세스) 동안만 유지, 재시작 시 초기화됨.


def _normalize_list(value):
    """API가 결과 1건일 때 dict, 여러 건일 때 list를 주는 것을 리스트로 통일한다."""
    if value is None:
        return []
    if isinstance(value, list):
        return value
    return [value]


def parse_search_response(data):
    """lawSearch.do 응답(JSON)을 파싱해 법령 목록을 반환한다."""
    law_search = data.get("LawSearch", {})
    laws = _normalize_list(law_search.get("law"))

    results = []
    for law in laws:
        results.append({
            "mst": law.get("법령일련번호", ""),
            "법령명": law.get("법령명한글", ""),
            "공포일자": law.get("공포일자", ""),
            "시행일자": law.get("시행일자", ""),
            "소관부처": law.get("소관부처명", ""),
            "법령구분": law.get("법령구분명", ""),
        })
    return results


def search_law(keyword, oc):
    """키워드로 법령을 검색해 목록을 반환한다. 네트워크/API 오류 시 예외를 던진다."""
    params = {"OC": oc, "target": "law", "type": "JSON", "query": keyword, "display": 20}
    resp = requests.get(f"{BASE_URL}/lawSearch.do", params=params, timeout=TIMEOUT)
    resp.raise_for_status()
    return parse_search_response(resp.json())


def parse_law_response(data):
    """lawService.do 응답(JSON)을 파싱해 법령 본문을 반환한다."""
    law = data.get("법령", {})
    basic = law.get("기본정보", {})
    articles_raw = _normalize_list(law.get("조문", {}).get("조문단위"))

    articles = []
    for art in articles_raw:
        articles.append({
            "조문번호": art.get("조문번호", ""),
            "조문제목": art.get("조문제목", ""),
            "조문내용": art.get("조문내용", ""),
        })

    return {
        "법령명": basic.get("법령명_한글", ""),
        "공포일자": basic.get("공포일자", ""),
        "시행일자": basic.get("시행일자", ""),
        "조문목록": articles,
    }


def get_law_text(mst, oc):
    """법령일련번호(mst)로 본문을 조회한다. 세션 캐시가 있으면 캐시를 반환한다."""
    if mst in _text_cache:
        return _text_cache[mst]

    params = {"OC": oc, "target": "law", "type": "JSON", "MST": mst}
    resp = requests.get(f"{BASE_URL}/lawService.do", params=params, timeout=TIMEOUT)
    resp.raise_for_status()
    result = parse_law_response(resp.json())
    _text_cache[mst] = result
    return result


def clear_cache():
    """세션 캐시를 비운다 (주로 테스트/재조회용)."""
    _text_cache.clear()
