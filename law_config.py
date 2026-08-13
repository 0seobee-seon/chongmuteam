"""
법률검토 모듈의 OC(국가법령정보 API 키) 로컬 저장/로드.
소스 코드에 OC 값을 하드코딩하지 않기 위한 계층.
"""

import json
import os

CONFIG_PATH = os.path.join(os.path.dirname(__file__), "config_법률검토.json")


def load_oc(config_path=CONFIG_PATH):
    """저장된 OC 키를 읽어온다. 파일이 없거나 손상되었으면 None을 반환한다."""
    if not os.path.exists(config_path):
        return None
    try:
        with open(config_path, "r", encoding="utf-8") as f:
            data = json.load(f)
        return data.get("oc") or None
    except (json.JSONDecodeError, OSError):
        return None


def save_oc(oc, config_path=CONFIG_PATH):
    """OC 키를 설정 파일에 저장한다."""
    with open(config_path, "w", encoding="utf-8") as f:
        json.dump({"oc": oc}, f, ensure_ascii=False)
