# Markdown to DOCX 변환기

마크다운 파일을 Word(DOCX) 포맷으로 변환합니다. 마크다운의 구조(제목, 목록, 굵은 글씨 등)를 완벽하게 보존합니다.

## 주요 기능

- ✅ **마크다운 파싱**: 제목(H1-H6), 목록, 인용문 자동 감지
- ✅ **인라인 포맷팅**: **bold**, *italic*, `code` 완벽 지원
- ✅ **HTML 마크업 처리**: `<br/>`, `<hr/>` 등 HTML 태그 자동 처리
- ✅ **통일된 폰트**: 기본값 맑은 고딕, 명령줄 옵션으로 변경 가능
- ✅ **구조 보존**: 마크다운 계층 구조 완벽하게 DOCX로 변환
- ✅ **다중 파일 처리**: 여러 마크다운 파일 일괄 변환 가능

## 설치

### 1. Node.js 설치 (v14 이상)
[nodejs.org](https://nodejs.org/)에서 다운로드 후 설치

### 2. 의존성 설치

```bash
npm install
```

설치되는 라이브러리:
- `docx`: DOCX 문서 생성 라이브러리

## 사용 방법

### 기본 사용 (맑은 고딕 폰트)

```bash
node markdown_to_docx.js
```
- 입력: `SKT_Anthropic_Meeting_251027.md`
- 출력: `SKT_Anthropic_Meeting_JS.docx`

### 다른 마크다운 파일 처리

```bash
node markdown_to_docx.js --input "파일명.md"
```

예시:
```bash
node markdown_to_docx.js --input "test_html.md"
```

### 폰트 변경

```bash
node markdown_to_docx.js --font "폰트명"
```

예시:
```bash
node markdown_to_docx.js --font "나눔고딕"
node markdown_to_docx.js --font "Arial"
node markdown_to_docx.js --font "Calibri"
```

### 옵션 조합

```bash
node markdown_to_docx.js --input "test.md" --font "나눔고딕"
```

## 지원하는 마크다운 문법

### 제목

```markdown
# Heading 1
## Heading 2
### Heading 3
#### Heading 4
##### Heading 5
###### Heading 6
```

### 강조

```markdown
**굵은 텍스트**     → Bold
*기울임 텍스트*     → Italic
`코드 블록`         → Code (Consolas 폰트, 배경색)
```

### 목록

```markdown
- 항목 1
- 항목 2
  - 중첩 항목 2-1
  - 중첩 항목 2-2
- 항목 3
```

### 인용문

```markdown
> 이것은 인용문입니다.
> 파란색 테두리와 회색 배경으로 표시됩니다.
```

### HTML 마크업

```markdown
첫 번째 줄<br/>두 번째 줄     → 줄바꿈 처리
텍스트<hr/>텍스트             → 구분선 처리
<!-- 주석 -->                  → 주석 제거
<div>...</div>                 → 태그 제거, 내용만 유지
```

## 출력 예시

### 입력 (Markdown)
```markdown
# SK x Anthropic 미팅

## 참석자

- **SK**: 최태원 (회장), 유영상 (CEO)
- **Anthropic**: Dario Amodei (CEO), Chris

## 주요 내용

한국 시장 진출 전략과 <br/> AI 협력 방안에 대해 논의했습니다.

> 양사의 파트너십은 한국의 디지털 혁신에 중요한 역할을 합니다.
```

### 출력 (DOCX)
```
Heading 1: SK x Anthropic 미팅

Heading 2: 참석자

• SK: 최태원 (회장), 유영상 (CEO)
• Anthropic: Dario Amodei (CEO), Chris

Heading 2: 주요 내용

일반 텍스트: 한국 시장 진출 전략과
일반 텍스트: AI 협력 방안에 대해 논의했습니다.

(인용문 - 파란색 테두리, 회색 배경)
양사의 파트너십은 한국의 디지털 혁신에 중요한 역할을 합니다.
```

## 스크립트 구조

```
markdown_to_docx.js
├─ processHtmlMarkup()        : HTML 태그 처리
├─ restoreHtmlMarkup()        : HTML 특수 마커 복원
├─ parseMarkdown()            : 마크다운 파싱
├─ parseInlineFormatting()    : 인라인 포맷팅 (bold, italic, code)
└─ Document 생성 및 저장
```

## 처리 단계

```
1. 명령줄 인자 파싱
   ├─ --font 옵션 처리
   └─ --input 옵션 처리

2. 마크다운 파일 읽기
   └─ UTF-8 인코딩

3. HTML 마크업 전처리
   ├─ <br/> → 특수 마커
   ├─ <hr/> → 특수 마커
   ├─ <!-- --> 제거
   └─ 기타 HTML 태그 제거

4. 마크다운 파싱
   ├─ 제목 감지 (#, ##, ###...)
   ├─ 목록 감지 (-, *)
   ├─ 인용문 감지 (>)
   └─ 일반 단락 처리

5. 인라인 포맷팅 처리
   ├─ **bold** 처리
   ├─ *italic* 처리
   ├─ `code` 처리
   └─ 특수 마커 복원

6. DOCX 문서 생성
   ├─ 스타일 적용
   ├─ 글꼴 통일 (맑은 고딕 또는 지정 폰트)
   └─ 파일 저장
```

## 출력 DOCX 파일

변환 후 생성되는 파일:

- **기본 입력 파일**: `SKT_Anthropic_Meeting_251027.md` → `SKT_Anthropic_Meeting_JS.docx`
- **--input 옵션**: `파일명.md` → `파일명_output.docx`

예시:
```bash
node markdown_to_docx.js --input "test.md"
# → test_output.docx 생성
```

## 글꼴 스타일

### 기본 적용

| 요소 | 폰트 | 크기 | 스타일 |
|------|------|------|--------|
| Heading 1 | 맑은 고딕 | 28pt | Bold |
| Heading 2 | 맑은 고딕 | 21pt | Bold |
| Heading 3 | 맑은 고딕 | 16pt | Bold |
| Heading 4 | 맑은 고딕 | 14pt | Bold |
| 본문 | 맑은 고딕 | 11pt | Regular |
| Code | Consolas | 10.5pt | Regular + 배경색 |

### 명령줄로 폰트 변경

```bash
# 나눔고딕 사용
node markdown_to_docx.js --font "나눔고딕"

# Arial 사용
node markdown_to_docx.js --font "Arial"
```

## 제한사항

1. **이미지**: 마크다운의 이미지 링크(`![alt](url)`)는 링크 텍스트만 유지
2. **테이블**: 마크다운 테이블은 일반 텍스트로 변환 (DOCX 테이블 미지원)
3. **링크**: URL은 파란색 텍스트로 표시되지만 하이퍼링크로 변환되지 않음
4. **폰트 가용성**: 지정한 폰트가 설치되지 않으면 시스템 기본 폰트로 대체

## 트러블슈팅

### "파일을 찾을 수 없습니다"

마크다운 파일이 동일한 디렉토리에 있는지 확인하세요.

```bash
# 파일 확인
ls *.md
```

### "font: 맑은 고딕"으로 설정했는데 다른 폰트로 표시됨

Windows에서 맑은 고딕이 설치되지 않았거나, Word에서 기본 폰트를 변경해야 할 수 있습니다.

```bash
# 설치된 폰트로 변경
node markdown_to_docx.js --font "Calibri"
```

### DOCX 파일이 열리지 않음

`docx` npm 패키지가 설치되었는지 확인하세요.

```bash
npm install docx
```

## 예제

### 예제 1: 기본 변환

```bash
node markdown_to_docx.js
```
결과: `SKT_Anthropic_Meeting_JS.docx` 생성

### 예제 2: 테스트 파일 변환

```bash
node markdown_to_docx.js --input "test_html.md"
```
결과: `test_html_output.docx` 생성

### 예제 3: 폰트 변경 + 파일 지정

```bash
node markdown_to_docx.js --input "test.md" --font "나눔고딕"
```
결과: `test_output.docx` 생성 (나눔고딕 폰트 적용)

## 패키지 정보

- **Node.js**: v14 이상 필요
- **주요 의존성**: `docx`
- **크기**: 약 5MB (node_modules 포함)

## 라이센스

ISC

## 참고 자료

- [docx npm 패키지](https://www.npmjs.com/package/docx)
- [마크다운 문법](https://www.markdownguide.org/)
- [Node.js 공식 사이트](https://nodejs.org/)
