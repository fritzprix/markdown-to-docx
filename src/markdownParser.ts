// src/markdownParser.ts
// 마크다운 파싱 로직

export interface TableElement {
  type: "table";
  headers: string[];
  rows: string[][];
  endIndex: number;
}

export interface CheckboxElement {
  type: "checkbox";
  level: number;
  text: string;
  checked: boolean;
}

export interface HeadingElement {
  type: "heading";
  level: number;
  text: string;
}

export interface ListElement {
  type: "list";
  level: number;
  text: string;
}

export interface BlockquoteElement {
  type: "blockquote";
  text: string;
}

export interface ParagraphElement {
  type: "paragraph";
  text: string;
}

export interface ImageElement {
  type: "image";
  alt: string; // alt 텍스트
  src: string; // 이미지 경로 또는 URL
  width?: number; // 선택적 너비 (포인트 단위)
  height?: number; // 선택적 높이 (포인트 단위)
}

export type ParsedElement =
  | TableElement
  | CheckboxElement
  | HeadingElement
  | ListElement
  | BlockquoteElement
  | ParagraphElement
  | ImageElement;

export interface ParseOptions {
  verbose?: boolean;
}

// HTML 마크업 처리 함수
export function processHtmlMarkup(text: string): string {
  // <br/>, <br>, <br /> 를 특수 문자로 임시 치환 (나중에 복원)
  text = text.replace(/<br\s*\/?>/gi, "\u0001BR\u0001");

  // <hr/>, <hr>, <hr /> 를 특수 문자로 임시 치환
  text = text.replace(/<hr\s*\/?>/gi, "\u0001HR\u0001");

  // HTML 주석 제거
  text = text.replace(/<!--[\s\S]*?-->/g, "");

  // 기타 HTML 태그 제거 (div, span, p 등)
  text = text.replace(/<[^>]+>/g, "");

  return text;
}

// 특수 마커를 복원하는 함수
export function restoreHtmlMarkup(text: string): string {
  // eslint-disable-next-line no-control-regex
  text = text.replace(/\u0001BR\u0001/g, "\n");
  // eslint-disable-next-line no-control-regex
  text = text.replace(/\u0001HR\u0001/g, "─────────────────────");
  return text;
}

// 마크다운 이미지 링크 파싱 함수 ![alt](src)
export function parseImageMarkdown(text: string): ImageElement | null {
  // ![alt text](image.png) 또는 ![alt](https://example.com/image.jpg) 형식
  // 선택적 너비/높이: ![alt](path){width=200,height=150}
  const imageRegex = /!\[([^\]]*)\]\(([^)]+)\)(?:{([^}]*)})?/;
  const match = text.match(imageRegex);

  if (!match) {
    return null;
  }

  const alt = match[1] || "Image";
  const src = match[2];
  const optionsStr = match[3];

  const image: ImageElement = {
    type: "image",
    alt,
    src,
  };

  // 선택적 너비/높이 파싱
  if (optionsStr) {
    const widthMatch = optionsStr.match(/width\s*=\s*(\d+)/);
    const heightMatch = optionsStr.match(/height\s*=\s*(\d+)/);

    if (widthMatch) {
      image.width = parseInt(widthMatch[1], 10);
    }

    if (heightMatch) {
      image.height = parseInt(heightMatch[1], 10);
    }
  }

  return image;
}

// 단락 텍스트에서 인라인 이미지 추출
export function extractInlineImages(text: string): ParsedElement[] {
  const elements: ParsedElement[] = [];
  const imageRegex = /!\[([^\]]*)\]\(([^)]+)\)/g;
  let lastIndex = 0;
  let match;

  // 정규표현식 exec 루프에서 할당이 필요함
  // noinspection JSAssignmentUsedAsCondition (WebStorm IDE comment)
  while ((match = imageRegex.exec(text)) !== null) {
    // 이미지 전의 텍스트
    if (match.index > lastIndex) {
      const beforeText = text.substring(lastIndex, match.index);
      if (beforeText.trim()) {
        elements.push({ type: "paragraph", text: beforeText });
      }
    }

    // 이미지 처리
    const alt = match[1] || "Image";
    const src = match[2];
    elements.push({
      type: "image",
      alt,
      src,
    } as ImageElement);

    lastIndex = match.index + match[0].length;
  }

  // 마지막 텍스트
  if (lastIndex < text.length) {
    const afterText = text.substring(lastIndex);
    if (afterText.trim()) {
      elements.push({ type: "paragraph", text: afterText });
    }
  }

  // 이미지가 없으면 원본 텍스트 반환
  if (elements.length === 0) {
    return [{ type: "paragraph", text }];
  }

  return elements;
}

// 테이블 파싱 함수 (| header | 형식)
export function parseTableMarkdown(lines: string[], startIndex: number): TableElement | null {
  const tableLines: string[] = [];
  let i = startIndex;

  // 첫 번째 줄이 테이블 헤더인지 확인
  if (!lines[i].includes("|")) {
    return null;
  }

  // 테이블의 모든 줄 수집
  while (i < lines.length && lines[i].trim().startsWith("|")) {
    tableLines.push(lines[i].trim());
    i++;
  }

  if (tableLines.length < 2) {
    return null;
  }

  // 헤더와 구분선 확인
  const headerLine = tableLines[0].split("|").slice(1, -1); // 양 끝 | 제거
  const separatorLine = tableLines[1];

  // 구분선이 올바른 형식인지 확인
  if (!separatorLine.includes("---")) {
    return null;
  }

  // 데이터 행들
  const rows = tableLines.slice(2).map((line) => line.split("|").slice(1, -1));

  return {
    type: "table",
    headers: headerLine.map((h) => h.trim()),
    rows: rows.map((row) => row.map((cell) => cell.trim())),
    endIndex: i,
  };
}

// 체크박스 파싱 함수 (- [ ] 또는 - [x] 형식)
export function isCheckbox(line: string): CheckboxElement | null {
  const checkboxRegex = /^(\s*)[-*]\s+\[[ xX]\]\s+(.+)$/;
  const match = line.match(checkboxRegex);

  if (match) {
    const checked = match[2].charAt(0) === "x" || match[2].charAt(0) === "X";
    const indent = match[1].length;
    const level = Math.floor(indent / 2);
    const text = match[2].substring(2).trim(); // [x] 제거

    return {
      type: "checkbox",
      level,
      text,
      checked,
    };
  }

  return null;
}

// 마크다운 파싱 함수
export function parseMarkdown(content: string, options: ParseOptions = {}): ParsedElement[] {
  const verbose = options.verbose !== false; // 기본값: true

  // HTML 마크업 처리 (특수 마커로 임시 치환)
  content = processHtmlMarkup(content);

  // Windows 줄바꿈 처리 (\r\n)
  const lines = content.replace(/\r\n/g, "\n").replace(/\r/g, "\n").split("\n");
  const elements: ParsedElement[] = [];
  let i = 0;

  while (i < lines.length) {
    let line = lines[i];

    // 특수 마커 복원
    line = restoreHtmlMarkup(line);

    // 빈 줄 스킵
    if (!line.trim()) {
      i++;
      continue;
    }

    // 테이블 처리
    const tableResult = parseTableMarkdown(lines, i);
    if (tableResult) {
      elements.push(tableResult);
      i = tableResult.endIndex;
      if (verbose)
        console.log(`Table: ${tableResult.headers.length} 열, ${tableResult.rows.length} 행`);
      continue;
    }

    // Heading 처리
    const headingMatch = line.match(/^(#{1,6})\s+(.+)$/);
    if (headingMatch) {
      const level = headingMatch[1].length;
      const text = headingMatch[2];
      elements.push({ type: "heading", level, text });
      if (verbose) console.log(`Heading ${level}: ${text.substring(0, 50)}`);
      i++;
      continue;
    }

    // 이미지 처리
    const imageResult = parseImageMarkdown(line);
    if (imageResult) {
      elements.push(imageResult);
      if (verbose) console.log(`Image: ${imageResult.alt} (${imageResult.src})`);
      i++;
      continue;
    }

    // 블록 인용 처리
    if (line.trim().startsWith(">")) {
      const quoteText = line.replace(/^>\s*/, "");
      elements.push({ type: "blockquote", text: quoteText });
      i++;
      continue;
    }

    // 체크박스 처리
    const checkboxResult = isCheckbox(line);
    if (checkboxResult) {
      elements.push(checkboxResult);
      i++;
      continue;
    }

    // 리스트 처리
    const listMatch = line.match(/^(\s*)[-*]\s+(.+)$/);
    if (listMatch) {
      const indent = listMatch[1].length;
      const level = Math.floor(indent / 2);
      const text = listMatch[2];
      elements.push({ type: "list", level, text });
      i++;
      continue;
    }

    // 일반 단락 (인라인 이미지 추출)
    const paragraphElements = extractInlineImages(line);
    elements.push(...paragraphElements);
    if (verbose && paragraphElements.some((e) => e.type === "image")) {
      paragraphElements.forEach((e) => {
        if (e.type === "image") {
          console.log(`Image (inline): ${(e as ImageElement).alt} (${(e as ImageElement).src})`);
        }
      });
    }
    i++;
  }

  if (verbose) {
    console.log(`\n총 ${elements.length}개 요소 파싱됨`);
    const headings = elements.filter((e) => e.type === "heading").length;
    const lists = elements.filter((e) => e.type === "list").length;
    const checkboxes = elements.filter((e) => e.type === "checkbox").length;
    const tables = elements.filter((e) => e.type === "table").length;
    const quotes = elements.filter((e) => e.type === "blockquote").length;
    const images = elements.filter((e) => e.type === "image").length;
    console.log(
      `Headings: ${headings}, Lists: ${lists}, Checkboxes: ${checkboxes}, Tables: ${tables}, Quotes: ${quotes}, Images: ${images}\n`
    );
  }

  return elements;
}
