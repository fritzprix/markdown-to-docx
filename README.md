# Markdown to DOCX 변환기

마크다운 파일을 Word(DOCX) 포맷으로 변환합니다. **라이브러리로도 사용 가능**하며 **CLI 도구**로도 사용할 수 있습니다. **TypeScript**로 작성되어 완벽한 타입 안정성을 제공합니다.

## 주요 기능

- ✅ **마크다운 파싱**: 제목(H1-H6), 목록, 인용문 자동 감지
- ✅ **인라인 포맷팅**: **bold**, _italic_, `code` 완벽 지원
- ✅ **테이블**: 마크다운 테이블 지원
- ✅ **링크**: `[텍스트](URL)` 형식 지원
- ✅ **이미지 링크**: `![alt](image.png)` 형식 지원
- ✅ **체크박스**: `- [ ]`, `- [x]` 형식 지원
- ✅ **HTML 마크업**: `<br/>`, `<hr/>` 등 HTML 태그 자동 처리
- ✅ **통일된 폰트**: 기본값 맑은 고딕, 명령줄/코드로 변경 가능
- ✅ **모듈화**: 라이브러리로 재사용 가능
- ✅ **TypeScript**: 완벽한 타입 안정성

## 설치

### 1. Node.js 설치 (v14 이상)

[nodejs.org](https://nodejs.org/)에서 다운로드 후 설치

### 2. 의존성 설치 및 빌드

```bash
npm install
npm run build
```

## 사용 방법

### CLI 사용 (TypeScript 컴파일 후)

#### 기본 사용 (input.md → output.docx)

```bash
npm start
# 또는
markdown-to-docx
```

#### 파일명 지정

```bash
npm start -- --input "마크다운파일.md"
# 또는
markdown-to-docx --input "마크다운파일.md"
```

#### 폰트 변경

```bash
npm start -- --font "나눔고딕"
# 또는
markdown-to-docx --font "나눔고딕"
```

#### 출력 파일명 지정

```bash
npm start -- --input "source.md" --output "result.docx"
# 또는
markdown-to-docx --input "source.md" --output "result.docx"
```

#### 옵션 조합

```bash
npm start -- --input "문서.md" --output "결과.docx" --font "Arial"
# 또는
markdown-to-docx --input "문서.md" --output "결과.docx" --font "Arial"
```

### 라이브러리 사용 (JavaScript)

#### 기본 사용

```javascript
const { convertMarkdownToDOCX } = require("./dist/index");

convertMarkdownToDOCX("input.md", "output.docx", {
  fontFamily: "맑은 고딕",
  verbose: true,
})
  .then(() => console.log("완료!"))
  .catch((error) => console.error(error));
```

#### 커스텀 빌더 사용

```javascript
const fs = require("fs");
const { parseMarkdown, DocxBuilder } = require("./dist/index");

// 마크다운 파싱
const content = fs.readFileSync("input.md", "utf-8");
const elements = parseMarkdown(content, { verbose: true });

// 커스텀 빌더 생성
const builder = new DocxBuilder("Arial");

// 요소 추가
elements.forEach((element) => {
  builder.addElement(element);
});

// 파일 저장
builder.save("output.docx");
```

### 라이브러리 사용 (TypeScript)

#### 기본 사용

```typescript
import { convertMarkdownToDOCX } from "./src/index";

async function main() {
  const outputFile = await convertMarkdownToDOCX("input.md", "output.docx", {
    fontFamily: "맑은 고딕",
    verbose: true,
  });
  console.log(`완료: ${outputFile}`);
}

main().catch(console.error);
```

#### 개발 중 ts-node로 직접 실행

```bash
npm run dev
```

또는

```bash
ts-node src/cli.ts --input "마크다운파일.md"
```

#### 커스텀 빌더 사용

```typescript
import { parseMarkdown, DocxBuilder } from "./src/index";
import * as fs from "fs";

async function main() {
  // 마크다운 파싱
  const content = fs.readFileSync("input.md", "utf-8");
  const elements = parseMarkdown(content, { verbose: true });

  // 커스텀 빌더 생성
  const builder = new DocxBuilder("Arial");

  // 요소 추가
  elements.forEach((element) => {
    builder.addElement(element);
  });

  // 파일 저장
  await builder.save("output.docx");
}

main().catch(console.error);
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
**굵은 텍스트** → Bold
_기울임 텍스트_ → Italic
`코드 블록` → Code (Consolas 폰트, 배경색)
```

### 목록

```markdown
- 항목 1
- 항목 2
  - 중첩 항목 2-1
  - 중첩 항목 2-2
- 항목 3
```

### 체크박스

```markdown
- [ ] 미완료 작업
- [x] 완료된 작업
```

### 테이블

```markdown
| 열1    | 열2    | 열3    |
| ------ | ------ | ------ |
| 데이터 | 데이터 | 데이터 |
| 데이터 | 데이터 | 데이터 |
```

### 링크

```markdown
[링크 텍스트](https://example.com)
```

### 이미지 링크

```markdown
![대체 텍스트](https://example.com/image.png)
```

### 인용문

```markdown
> 이것은 인용문입니다.
> 파란색 테두리와 회색 배경으로 표시됩니다.
```

### HTML 마크업

```markdown
첫 번째 줄<br/>두 번째 줄 → 줄바꿈 처리
텍스트<hr/>텍스트 → 구분선 처리

<!-- 주석 -->                  → 주석 제거
<div>...</div>                 → 태그 제거, 내용만 유지
```

## 프로젝트 구조

```
markdown-to-docx/
├── src/                          # TypeScript 소스
│   ├── index.ts                  # 메인 진입점 & 라이브러리 API
│   ├── cli.ts                    # CLI 진입점
│   ├── markdownParser.ts         # 마크다운 파싱
│   ├── inlineFormatter.ts        # 인라인 포맷팅 (bold, italic, link 등)
│   └── docxBuilder.ts            # DOCX 문서 생성
├── dist/                         # 컴파일된 JavaScript (npm run build로 생성)
│   ├── index.js, index.d.ts
│   ├── cli.js, cli.d.ts
│   ├── markdownParser.js, markdownParser.d.ts
│   ├── inlineFormatter.js, inlineFormatter.d.ts
│   └── docxBuilder.js, docxBuilder.d.ts
├── examples/                     # 사용 예제
│   ├── basic-usage.js            # JavaScript 기본 사용 예제
│   ├── advanced-usage.js         # JavaScript 고급 사용 예제
│   └── ts-usage.ts               # TypeScript 사용 예제
├── lib/                          # 레거시 JavaScript 소스 (더 이상 사용 안 함)
├── tsconfig.json                 # TypeScript 설정
├── package.json                  # 프로젝트 설정 & npm 스크립트
├── README.md                     # 이 파일
└── .gitignore
```

## API 레퍼런스

### convertMarkdownToDOCX(inputFile, outputFile, options)

마크다운 파일을 DOCX로 변환합니다.

**매개변수:**

- `inputFile` (string): 입력 마크다운 파일 경로
- `outputFile` (string): 출력 DOCX 파일 경로
- `options` (object): 옵션
  - `fontFamily` (string): 사용할 폰트 (기본값: '맑은 고딕')
  - `verbose` (boolean): 상세 로그 출력 (기본값: true)

**반환값:**

- `Promise<string>`: 생성된 파일 경로

**예제:**

```javascript
await convertMarkdownToDOCX("input.md", "output.docx", {
  fontFamily: "Arial",
  verbose: true,
});
```

### parseMarkdown(content, options)

마크다운 텍스트를 파싱합니다.

**매개변수:**

- `content` (string): 마크다운 텍스트
- `options` (object): 옵션
  - `verbose` (boolean): 상세 로그 출력 (기본값: true)

**반환값:**

- `Array<Object>`: 파싱된 요소 배열

**예제:**

```javascript
const elements = parseMarkdown(markdownText, { verbose: false });
```

### DocxBuilder 클래스

DOCX 문서를 생성합니다.

**메서드:**

- `addElement(element)`: 요소 추가
- `addHeading(element)`: 제목 추가
- `addTable(element)`: 테이블 추가
- `addBlockquote(element)`: 인용문 추가
- `addCheckbox(element)`: 체크박스 추가
- `addList(element)`: 목록 추가
- `addParagraph(element)`: 단락 추가
- `build()`: 문서 생성
- `save(outputFile)`: 파일 저장

**예제:**

```javascript
const builder = new DocxBuilder("Arial");
elements.forEach((el) => builder.addElement(el));
await builder.save("output.docx");
```

## 트러블슈팅

### "파일을 찾을 수 없습니다"

마크다운 파일이 동일한 디렉토리에 있는지 확인하세요.

```bash
# 파일 확인
dir *.md  (또는 ls *.md)

# 파일명 지정
node cli.js --input "정확한파일명.md"
```

### DOCX 파일이 열리지 않음

`docx` npm 패키지가 설치되었는지 확인하세요.

```bash
npm install docx
```

### 폰트가 적용되지 않음

Word에서 설정한 기본 폰트가 다를 수 있습니다. 문서를 연 후 "모두 선택(Ctrl+A)"으로 폰트를 변경하세요.

## 예제

### 예제 1: 기본 CLI 사용

```bash
node cli.js --input "meeting_notes.md"
```

### 예제 2: 라이브러리 기본 사용

```bash
node examples/basic-usage.js
```

### 예제 3: 라이브러리 고급 사용

```bash
node examples/advanced-usage.js
```

## 라이선스

ISC

## 참고 자료

- [docx npm 패키지](https://www.npmjs.com/package/docx)
- [마크다운 문법](https://www.markdownguide.org/)
- [Node.js 공식 사이트](https://nodejs.org/)
