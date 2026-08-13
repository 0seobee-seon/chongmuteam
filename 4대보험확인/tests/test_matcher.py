import sys
import os
sys.path.insert(0, os.path.join(os.path.dirname(__file__), ".."))

from matcher import classify_errors, merge_records


def _payroll(사원번호, 성명, 생년월일6자리, **overrides):
    record = {
        "사원번호": 사원번호, "성명": 성명, "생년월일6자리": 생년월일6자리,
        "성별코드": "1", "부서": "총무팀", "퇴사일": None,
        "국민연금_급여대장": 100000, "건강보험_급여대장": 200000,
        "고용보험_급여대장": 30000, "장기요양보험_급여대장": 20000,
    }
    record.update(overrides)
    return record


def test_merge_matches_by_name_and_birthdate():
    payroll = [_payroll("9001", "김테스트", "900101")]
    health = [{"성명": "김테스트", "생년월일6자리": "900101", "성별코드": "1",
               "건강보험_공단": 200000, "장기요양보험_공단": 20000, "상실일": 99991231}]
    pension = [{"성명": "김테스트", "생년월일6자리": "900101", "성별코드": "1",
                "국민연금_공단": 100000}]
    employment = [{"성명": "김테스트", "생년월일6자리": "900101",
                   "고용보험_공단": 30000, "고용종료일": None}]

    merged = merge_records(payroll, health, pension, employment)

    assert len(merged) == 1
    assert merged[0]["건강보험_공단"] == 200000
    assert merged[0]["국민연금_공단"] == 100000
    assert merged[0]["고용보험_공단"] == 30000
    assert merged[0]["확인필요"] is False


def test_merge_leaves_none_when_no_match_in_corp_data():
    payroll = [_payroll("9002", "이가상", "900101")]

    merged = merge_records(payroll, [], [], [])

    assert merged[0]["건강보험_공단"] is None
    assert merged[0]["국민연금_공단"] is None
    assert merged[0]["고용보험_공단"] is None


def test_merge_flags_duplicate_key_as_확인필요():
    payroll = [
        _payroll("9003", "박동명", "800101"),
        _payroll("9099", "박동명", "800101"),
    ]

    merged = merge_records(payroll, [], [], [])

    assert merged[0]["확인필요"] is True
    assert merged[1]["확인필요"] is True


def _merged(**overrides):
    record = {
        "사원번호": "9001", "성명": "김테스트", "생년월일6자리": "900101",
        "부서": "총무팀", "퇴사일": None, "확인필요": False,
        "국민연금_급여대장": 100000, "국민연금_공단": 100000,
        "건강보험_급여대장": 200000, "건강보험_공단": 200000, "건강보험_상실일": 99991231,
        "장기요양보험_급여대장": 20000, "장기요양보험_공단": 20000,
        "고용보험_급여대장": 30000, "고용보험_공단": 30000, "고용보험_고용종료일": None,
    }
    record.update(overrides)
    return record


def test_classify_errors_no_error_when_all_match():
    errors = classify_errors(_merged(), "202607")
    assert errors == []


def test_classify_errors_amount_mismatch():
    record = _merged(건강보험_급여대장=200000, 건강보험_공단=190000)
    errors = classify_errors(record, "202607")

    assert len(errors) == 1
    assert errors[0]["오류유형"] == "금액 불일치"
    assert errors[0]["보험종류"] == "건강보험"
    assert errors[0]["차액"] == 10000


def test_classify_errors_missing_in_corp_data():
    record = _merged(고용보험_급여대장=30000, 고용보험_공단=None)
    errors = classify_errors(record, "202607")

    assert len(errors) == 1
    assert errors[0]["오류유형"] == "공단에만 없음"
    assert errors[0]["보험종류"] == "고용보험"


def test_classify_errors_missing_in_payroll():
    record = _merged(국민연금_급여대장=None, 국민연금_공단=100000)
    errors = classify_errors(record, "202607")

    assert len(errors) == 1
    assert errors[0]["오류유형"] == "급여대장에만 없음"


def test_classify_errors_corp_zero_amount_and_payroll_blank_is_not_an_error():
    # 공단 금액이 0원(부과 없음)이고 급여대장은 공란인 경우 — 둘 다 "부과 없음"을 뜻하므로 정상.
    record = _merged(건강보험_급여대장=None, 건강보험_공단=0)
    errors = classify_errors(record, "202607")

    types = [e["보험종류"] for e in errors]
    assert "건강보험" not in types


def test_classify_errors_payroll_zero_amount_and_corp_blank_is_not_an_error():
    record = _merged(건강보험_급여대장=0, 건강보험_공단=None)
    errors = classify_errors(record, "202607")

    types = [e["보험종류"] for e in errors]
    assert "건강보험" not in types


def test_classify_errors_corp_zero_but_payroll_has_real_amount_is_still_an_error():
    record = _merged(건강보험_급여대장=200000, 건강보험_공단=0)
    errors = classify_errors(record, "202607")

    assert len(errors) == 1
    assert errors[0]["오류유형"] == "공단에만 없음"


def test_classify_errors_both_blank_is_not_enrolled_no_error():
    record = _merged(국민연금_급여대장=None, 국민연금_공단=None)
    errors = classify_errors(record, "202607")
    assert errors == []


def test_classify_errors_terminated_employee_still_charged():
    record = _merged(건강보험_상실일=20260601)  # 급여 귀속월(202607)보다 이전 상실
    errors = classify_errors(record, "202607")

    types = [e["오류유형"] for e in errors]
    assert "전월 퇴사자 여전히 부과" in types


def test_classify_errors_active_employee_sentinel_date_is_not_terminated():
    record = _merged(건강보험_상실일=99991231)
    errors = classify_errors(record, "202607")

    types = [e["오류유형"] for e in errors]
    assert "전월 퇴사자 여전히 부과" not in types


def test_classify_errors_only_checks_requested_items():
    # 국민연금 자료가 아예 제공되지 않은 상황을 흉내낸다(공단금액=None인데 급여대장엔 금액 있음).
    # items에서 국민연금을 빼면, 파일이 없어서 못 비교한 것이므로 오류로 잡히면 안 된다.
    record = _merged(국민연금_급여대장=100000, 국민연금_공단=None)
    errors = classify_errors(record, "202607", items=["건강보험", "장기요양보험", "고용보험"])

    types = [e["보험종류"] for e in errors]
    assert "국민연금" not in types


def test_classify_errors_still_checks_terminated_flag_only_for_requested_items():
    record = _merged(건강보험_상실일=20260601)
    errors = classify_errors(record, "202607", items=["국민연금"])

    types = [e["오류유형"] for e in errors]
    assert "전월 퇴사자 여전히 부과" not in types
