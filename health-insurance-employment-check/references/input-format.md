# 비교 스크립트 입력 형식

UTF-8 JSON 파일을 사용한다.

```json
{
  "people": [
    {
      "name": "홍길동",
      "workbook": {"organization": "예시기관", "date": "2025-10-30", "career_type": "취업"},
      "insurance_records": [
        {"subscriber_type": "직장가입자", "organization": "예시기관", "acquired_date": "2025-10-30", "lost_date": null}
      ]
    }
  ]
}
```

날짜는 가능한 한 `YYYY-MM-DD` 문자열로 전달한다. 값이 판독되지 않으면 추측하지 말고 `null`을 사용한다. 스크립트는 표준 출력으로 JSON 결과를 내보낸다.
