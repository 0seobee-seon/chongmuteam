"""4대보험 확인 프로그램 — 매칭 및 오류 판정 모듈."""

import re

INSURANCE_ITEMS = ["국민연금", "건강보험", "장기요양보험", "고용보험"]

_TERMINATION_DATE_FIELDS = {
    "건강보험": "건강보험_상실일",
    "장기요양보험": "건강보험_상실일",
    "고용보험": "고용보험_고용종료일",
}


def _build_key(성명, 생년월일6자리):
    return (str(성명).strip(), 생년월일6자리)


def _index_by_key(records):
    result = {}
    for record in records:
        key = _build_key(record["성명"], record["생년월일6자리"])
        result[key] = record
    return result


def _find_duplicate_keys(payroll_records):
    seen = set()
    duplicates = set()
    for record in payroll_records:
        key = _build_key(record["성명"], record["생년월일6자리"])
        if key in seen:
            duplicates.add(key)
        seen.add(key)
    return duplicates


def merge_records(payroll_records, health_records, pension_records, employment_records):
    """급여대장 레코드를 기준으로 공단 3개 자료를 성명+생년월일6자리 키로 매칭한다."""
    health_by_key = _index_by_key(health_records)
    pension_by_key = _index_by_key(pension_records)
    employment_by_key = _index_by_key(employment_records)
    duplicate_keys = _find_duplicate_keys(payroll_records)

    merged = []
    for payroll in payroll_records:
        key = _build_key(payroll["성명"], payroll["생년월일6자리"])
        health = health_by_key.get(key)
        pension = pension_by_key.get(key)
        employment = employment_by_key.get(key)

        merged.append({
            "사원번호": payroll["사원번호"],
            "성명": payroll["성명"],
            "생년월일6자리": payroll["생년월일6자리"],
            "부서": payroll["부서"],
            "퇴사일": payroll["퇴사일"],
            "확인필요": key in duplicate_keys,
            "국민연금_급여대장": payroll["국민연금_급여대장"],
            "국민연금_공단": pension["국민연금_공단"] if pension else None,
            "건강보험_급여대장": payroll["건강보험_급여대장"],
            "건강보험_공단": health["건강보험_공단"] if health else None,
            "건강보험_상실일": health["상실일"] if health else None,
            "장기요양보험_급여대장": payroll["장기요양보험_급여대장"],
            "장기요양보험_공단": health["장기요양보험_공단"] if health else None,
            "고용보험_급여대장": payroll["고용보험_급여대장"],
            "고용보험_공단": employment["고용보험_공단"] if employment else None,
            "고용보험_고용종료일": employment["고용종료일"] if employment else None,
        })
    return merged


def classify_errors(merged_record, 귀속년월, items=None):
    """통합 레코드 1건에 대해 보험 항목별 오류를 판정한다.

    items를 지정하면 그 항목만 판정한다 — 공단 파일을 아예 선택하지 않아
    비교 자체가 불가능한 항목을 오류로 오탐하지 않기 위함(예: 국민연금 자료가
    아직 없어 선택하지 않은 달). 기본값은 4개 항목 전체.
    """
    if items is None:
        items = INSURANCE_ITEMS

    errors = []

    for item in items:
        급여대장금액 = merged_record.get(f"{item}_급여대장")
        공단금액 = merged_record.get(f"{item}_공단")
        # 0원 부과는 "부과 없음"과 같은 뜻이므로 공란(None)과 동일하게 취급한다.
        급여대장_없음 = 급여대장금액 is None or 급여대장금액 == 0
        공단_없음 = 공단금액 is None or 공단금액 == 0

        if 급여대장_없음 and 공단_없음:
            continue
        if not 급여대장_없음 and 공단_없음:
            errors.append(_make_error(merged_record, item, 급여대장금액, 공단금액, "공단에만 없음"))
        elif 급여대장_없음 and not 공단_없음:
            errors.append(_make_error(merged_record, item, 급여대장금액, 공단금액, "급여대장에만 없음"))
        elif 급여대장금액 != 공단금액:
            errors.append(_make_error(merged_record, item, 급여대장금액, 공단금액, "금액 불일치"))

    for item, date_field in _TERMINATION_DATE_FIELDS.items():
        if item not in items:
            continue
        상실일 = merged_record.get(date_field)
        공단금액 = merged_record.get(f"{item}_공단")
        if 공단금액 and _is_terminated_before(상실일, 귀속년월):
            errors.append(_make_error(
                merged_record, item, merged_record.get(f"{item}_급여대장"),
                공단금액, "전월 퇴사자 여전히 부과",
            ))

    return errors


def _make_error(record, item, 급여대장금액, 공단금액, 오류유형):
    차액 = None
    if 급여대장금액 is not None and 공단금액 is not None:
        차액 = 급여대장금액 - 공단금액
    return {
        "사원번호": record["사원번호"],
        "성명": record["성명"],
        "부서": record["부서"],
        "보험종류": item,
        "급여대장금액": 급여대장금액,
        "공단금액": 공단금액,
        "차액": 차액,
        "오류유형": 오류유형,
    }


def _is_terminated_before(상실일, 귀속년월):
    if not 상실일 or not 귀속년월:
        return False
    상실일_yyyymmdd = _to_yyyymmdd(상실일)
    if 상실일_yyyymmdd is None or 상실일_yyyymmdd == "99991231":
        return False
    return 상실일_yyyymmdd[:6] < 귀속년월


def _to_yyyymmdd(value):
    if hasattr(value, "strftime"):
        return value.strftime("%Y%m%d")
    if isinstance(value, (int, float)):
        return str(int(value))
    if isinstance(value, str):
        digits = re.sub(r"\D", "", value)
        return digits if len(digits) == 8 else None
    return None
