# 🤖 AI Test Case Generator

요구사항 문서(PDF/PPTX/DOCX)를 업로드하면 **Claude AI**가 자동으로 Test Case Excel을 생성해주는 웹앱입니다.

## ✨ 주요 기능

- 📄 PDF, PowerPoint, Word, 텍스트 파일 지원
- 🤖 Claude AI 기반 요구사항 자동 분석
- 📊 카테고리별 시트 분리 Excel 생성
- 🎯 필수/선택/개발 요구사항 자동 분류
- ⬇️ 바로 다운로드 가능

## 🚀 실행 방법

### 로컬 실행
```bash
pip install -r requirements.txt
streamlit run app.py
```

### Streamlit Cloud 배포
1. 이 Repository를 Fork
2. [streamlit.io](https://streamlit.io) 에서 GitHub 연동
3. 자동 배포 완료 → 영구 URL 발급

## 🔑 API Key 설정

[Anthropic Console](https://console.anthropic.com) 에서 API Key 발급 후 앱 사이드바에 입력하세요.

## 📋 생성되는 TC 항목

| 컬럼 | 내용 |
|------|------|
| TC ID | 고유 테스트 케이스 번호 |
| 분류 | 기능 카테고리 |
| 요구사항 유형 | 필수/선택/개발 |
| 테스트 항목 | TC 제목 |
| 테스트 절차 | 구체적인 테스트 방법 |
| 기대 결과 | 통과 기준 |
| 우선순위 | High/Medium/Low |
| 결과(Pass/Fail) | 테스트 결과 기입란 |

## 📁 프로젝트 구조

```
tc_app/
├── app.py              # 메인 Streamlit 앱
├── requirements.txt    # 패키지 목록
└── README.md           # 이 파일
```

---
Made with ❤️ using Claude AI & Streamlit
