#!/usr/bin/env python3
import argparse
import json
import re
import sys
import unicodedata
from datetime import date, datetime, timedelta


def normalize_org(value):
    if value is None:
        return None
    text = unicodedata.normalize("NFKC", str(value)).strip().casefold()
    text = re.sub(r"^(?:주식회사|\(주\)|주\))", "", text)
    text = re.sub(r"(?:주식회사|\(주\)|주\))$", "", text)
    return re.sub(r"[\s\-_.·,()]+", "", text) or None


def normalize_date(value):
    if value in (None, ""):
        return None
    if isinstance(value, (int, float)):
        return (date(1899, 12, 30) + timedelta(days=float(value))).isoformat()
    text = unicodedata.normalize("NFKC", str(value)).strip().rstrip(".")
    for fmt in ("%Y-%m-%d", "%Y/%m/%d", "%Y.%m.%d", "%Y%m%d"):
        try:
            return datetime.strptime(text, fmt).date().isoformat()
        except ValueError:
            pass
    return None


def make_result(person, workbook, record, status, note):
    record = record or {}
    return {
        "name": person.get("name"),
        "workbook_organization": workbook.get("organization"),
        "insurance_organization": record.get("organization"),
        "workbook_date": normalize_date(workbook.get("date")),
        "insurance_acquired_date": normalize_date(record.get("acquired_date")),
        "organization_match": (
            normalize_org(workbook.get("organization")) == normalize_org(record.get("organization"))
            if record else None
        ),
        "date_match": (
            normalize_date(workbook.get("date")) == normalize_date(record.get("acquired_date"))
            if record else None
        ),
        "status": status,
        "note": note,
    }


def compare_person(person):
    workbook = person.get("workbook") or {}
    expected_org = normalize_org(workbook.get("organization"))
    expected_date = normalize_date(workbook.get("date"))
    records = [
        record for record in (person.get("insurance_records") or [])
        if record.get("subscriber_type") == "직장가입자"
    ]

    if workbook.get("career_type") == "창업":
        return make_result(person, workbook, None, "확인 필요", "창업은 건강보험 직장가입 이력만으로 입증할 수 없음")
    if not expected_org or not expected_date or not records:
        return make_result(person, workbook, None, "확인 필요", "비교할 기관명·일자 또는 직장가입 이력이 없음")

    scored = []
    for record in records:
        org_match = normalize_org(record.get("organization")) == expected_org
        date_match = normalize_date(record.get("acquired_date")) == expected_date
        scored.append((int(org_match) + int(date_match), org_match, date_match, record))
    scored.sort(key=lambda item: item[0], reverse=True)
    best = scored[0]
    if sum(item[0] == best[0] for item in scored) > 1 and best[0] < 2:
        return make_result(person, workbook, None, "확인 필요", "동점인 직장가입 이력이 여러 개임")

    _, org_match, date_match, record = best
    if org_match and date_match:
        status = "일치"
    elif date_match:
        status = "기관명 불일치"
    elif org_match:
        status = "날짜 불일치"
    else:
        status = "기관명·날짜 불일치"
    return make_result(person, workbook, record, status, "")


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("input", help="UTF-8 JSON input file")
    args = parser.parse_args()
    with open(args.input, encoding="utf-8") as stream:
        payload = json.load(stream)
    output = {"results": [compare_person(person) for person in payload.get("people", [])]}
    json.dump(output, sys.stdout, ensure_ascii=False, indent=2)
    sys.stdout.write("\n")


if __name__ == "__main__":
    main()
