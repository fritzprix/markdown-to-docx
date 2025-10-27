# 마크다운 → DOCX 변환기 (TypeScript)

## 프로젝트 개요

이 프로젝트는 마크다운 파일을 Word(DOCX) 문서로 변환하는 도구입니다. **TypeScript**로 완벽하게 작성되어 있으며, 라이브러리와 CLI 도구 모두로 사용 가능합니다.

## 주요 특징

✅ **완전한 마크다운 지원**

- 제목 (H1-H6), 목록, 인용문
- 테이블, 체크박스
- 인라인 포맷팅 (**bold**, _italic_, `code`)
- 링크 및 이미지 링크

✅ **TypeScript + 타입 안정성**

- 완벽한 타입 정의 및 인터페이스
- 런타임 에러 방지

✅ **개발 환경**

- ESLint 기반 코드 검사
- Prettier 기반 자동 포맷팅
- npm 스크립트로 간편한 실행

✅ **배포 준비**

- npm 전역 설치 지원
- npx 실행 가능
- package.json bin 필드 설정

## 빠른 시작

### 설치

```bash
npm install
```

### 빌드

```bash
npm run build
```

### CLI 사용

```bash
# 기본 사용 (input.md → output.docx)
npm start

# 파일 지정
npm start -- --input "파일.md" --output "결과.docx"

# 또는 전역 명령어 (npm link 후)
markdown-to-docx --input "파일.md"
```

### TypeScript로 직접 실행

```bash
npx ts-node src/cli.ts --input "파일.md"
```

### 라이브러리로 사용

```typescript
import { convertMarkdownToDOCX, DocxBuilder, parseMarkdown } from "./src/index";

// 기본 사용
await convertMarkdownToDOCX("input.md", "output.docx", {
  fontFamily: "맑은 고딕",
});
```

## npm 스크립트

| 스크립트               | 설명                          |
| ---------------------- | ----------------------------- |
| `npm run build`        | TypeScript 컴파일 (dist 생성) |
| `npm start`            | 빌드 후 CLI 실행              |
| `npm run dev`          | ts-node로 개발 모드 실행      |
| `npm run lint`         | ESLint로 코드 검사            |
| `npm run lint:fix`     | ESLint로 자동 수정            |
| `npm run format`       | Prettier로 포맷팅             |
| `npm run format:check` | 포맷팅 확인 (수정 없음)       |

## 프로젝트 구조

```
.
├── src/                          # TypeScript 소스
│   ├── cli.ts                    # CLI 진입점
│   ├── index.ts                  # 라이브러리 API
│   ├── markdownParser.ts         # 마크다운 파싱
│   ├── inlineFormatter.ts        # 인라인 포맷팅
│   └── docxBuilder.ts            # DOCX 생성
├── dist/                         # 컴파일된 JavaScript (git 제외)
├── examples/                     # 사용 예제
│   ├── basic-usage.ts            # 기본 예제
│   ├── advanced-usage.ts         # 고급 예제
│   └── ts-usage.ts               # TypeScript 예제
├── tests/                        # 테스트 파일
│   ├── test_features.md
│   ├── test_formatting.md
│   └── test_html.md
├── build/                        # 생성된 DOCX 파일 (git 제외)
├── .eslintrc.json                # ESLint 설정 (구형)
├── eslint.config.cjs             # ESLint 설정 (현재)
├── .prettierrc                   # Prettier 설정
├── .prettierignore               # Prettier 제외 파일
├── tsconfig.json                 # TypeScript 설정
├── package.json                  # 프로젝트 설정
└── README.md                     # 문서
```

## API 레퍼런스

### convertMarkdownToDOCX(inputFile, outputFile, options)

마크다운 파일을 DOCX로 변환합니다.

**매개변수:**

- `inputFile` (string): 입력 마크다운 파일 경로
- `outputFile` (string): 출력 DOCX 파일 경로
- `options` (object):
  - `fontFamily` (string): 사용할 폰트 (기본값: '맑은 고딕')
  - `verbose` (boolean): 로그 출력 (기본값: true)

**반환값:** Promise<string> - 생성된 파일 경로

### parseMarkdown(content, options)

마크다운 텍스트를 구조화된 요소로 파싱합니다.

**매개변수:**

- `content` (string): 마크다운 텍스트
- `options` (object):
  - `verbose` (boolean): 로그 출력 (기본값: true)

**반환값:** ParsedElement[] - 파싱된 요소 배열

### DocxBuilder 클래스

DOCX 문서를 생성하는 빌더 클래스입니다.

**메서드:**

- `addElement(element)`: 요소 추가
- `addHeading(element)`: 제목 추가
- `addTable(element)`: 테이블 추가
- `addBlockquote(element)`: 인용문 추가
- `addCheckbox(element)`: 체크박스 추가
- `addList(element)`: 목록 추가
- `addParagraph(element)`: 단락 추가
- `build()`: Document 객체 생성
- `save(outputFile)`: 파일로 저장

## 개발 가이드

### 코드 스타일

- **포맷팅:** Prettier (.prettierrc)
- **린트:** ESLint (eslint.config.cjs)
- **언어:** TypeScript (strict mode)

### 커밋 전 체크리스트

```bash
# 1. 린트 확인
npm run lint

# 2. 포맷팅 확인
npm run format:check

# 3. 빌드 테스트
npm run build

# 4. 기능 테스트
npx ts-node src/cli.ts --input tests/test_features.md
```

## 지원하는 마크다운

- **제목:** `# H1` ~ `###### H6`
- **강조:** `**bold**`, `*italic*`, `` `code` ``
- **목록:** `- item` (중첩 지원)
- **체크박스:** `- [ ] todo`, `- [x] done`
- **테이블:** 표준 마크다운 테이블
- **링크:** `[text](url)`
- **이미지:** `![alt](url)`
- **인용문:** `> quote`
- **구분선:** `---` or `<hr/>`
- **개행:** `<br/>`

## 문제 해결

### npm link 문제

```bash
# npm link 제거
npm unlink markdown-to-docx -g

# 재설정
npm link
```

### 타입 에러

```bash
# TypeScript 재컴파일
npm run build
```

### 포맷팅 에러

```bash
# 자동 포맷팅
npm run format
```

## 라이센스

ISC

## 기여

버그 리포트 및 기능 요청은 Git 이슈로 등록해주세요.
