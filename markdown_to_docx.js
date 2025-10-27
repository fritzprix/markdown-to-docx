const fs = require("fs");
const {
  Document,
  Packer,
  Paragraph,
  TextRun,
  HeadingLevel,
  AlignmentType,
  LevelFormat,
  Table,
  TableCell,
  TableRow,
  BorderStyle,
  VerticalAlign,
} = require("docx");

// 명령줄 인자 처리
const args = process.argv.slice(2);
let fontFamily = "맑은 고딕"; // 기본 폰트
let inputFile = "input.md"; // 기본 입력 파일명
let outputFile = "output.docx";

// 인자 파싱
for (let i = 0; i < args.length; i++) {
  if (args[i] === "--font" && args[i + 1]) {
    fontFamily = args[i + 1];
  } else if (args[i] === "--input" && args[i + 1]) {
    inputFile = args[i + 1];
    outputFile = inputFile.replace(".md", "_output.docx");
  }
}

console.log(`사용 폰트: ${fontFamily}`);
console.log(`입력: ${inputFile}\n`);

// 마크다운 파일 읽기 (에러 처리)
let markdownContent;
try {
  if (!fs.existsSync(inputFile)) {
    throw new Error(`파일을 찾을 수 없습니다: ${inputFile}`);
  }
  markdownContent = fs.readFileSync(inputFile, "utf-8");
} catch (error) {
  console.error("\n❌ 오류 발생:");
  console.error(`\n  ${error.message}\n`);
  console.error("💡 해결 방법:");
  console.error(`  1. 파일이 현재 디렉토리에 있는지 확인하세요.`);
  console.error(`     명령어: dir *.md  (또는 ls *.md)\n`);
  console.error(`  2. 파일명을 정확히 지정하세요.`);
  console.error(`     명령어: node markdown_to_docx.js --input "파일명.md"\n`);
  console.error(`  3. 파일이 UTF-8 인코딩인지 확인하세요.\n`);
  console.error("📝 사용법:");
  console.error(`  기본 (input.md 찾기):`);
  console.error(`    node markdown_to_docx.js\n`);
  console.error(`  특정 파일 지정:`);
  console.error(`    node markdown_to_docx.js --input "마크다운파일.md"\n`);
  console.error(`  폰트 변경:`);
  console.error(`    node markdown_to_docx.js --font "나눔고딕"\n`);
  process.exit(1);
}

// HTML 마크업 처리 함수
function processHtmlMarkup(text) {
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
function restoreHtmlMarkup(text) {
  text = text.replace(/\u0001BR\u0001/g, "\n");
  text = text.replace(/\u0001HR\u0001/g, "─────────────────────");
  return text;
}

// 테이블 파싱 함수 (| header | 형식)
function parseTableMarkdown(lines, startIndex) {
  const tableLines = [];
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
  const rows = tableLines.slice(2).map((line) =>
    line.split("|").slice(1, -1)
  );

  return {
    type: "table",
    headers: headerLine.map((h) => h.trim()),
    rows: rows.map((row) => row.map((cell) => cell.trim())),
    endIndex: i,
  };
}

// 링크 파싱 함수 ([텍스트](URL) 형식)
function extractLinks(text) {
  const linkRegex = /\[([^\]]+)\]\(([^)]+)\)/g;
  const links = [];
  let match;

  while ((match = linkRegex.exec(text)) !== null) {
    links.push({
      text: match[1],
      url: match[2],
      fullMatch: match[0],
    });
  }

  return links;
}

// 체크박스 파싱 함수 (- [ ] 또는 - [x] 형식)
function isCheckbox(line) {
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
function parseMarkdown(content) {
  // HTML 마크업 처리 (특수 마커로 임시 치환)
  content = processHtmlMarkup(content);

  // Windows 줄바꿈 처리 (\r\n)
  const lines = content.replace(/\r\n/g, "\n").replace(/\r/g, "\n").split("\n");
  const elements = [];
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
      console.log(`Table: ${tableResult.headers.length} 열, ${tableResult.rows.length} 행`);
      continue;
    }

    // Heading 처리
    const headingMatch = line.match(/^(#{1,6})\s+(.+)$/);
    if (headingMatch) {
      const level = headingMatch[1].length;
      const text = headingMatch[2];
      elements.push({ type: "heading", level, text });
      console.log(`Heading ${level}: ${text.substring(0, 50)}`);
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

    // 일반 단락
    elements.push({ type: "paragraph", text: line });
    i++;
  }

  console.log(`\n총 ${elements.length}개 요소 파싱됨`);
  const headings = elements.filter((e) => e.type === "heading").length;
  const lists = elements.filter((e) => e.type === "list").length;
  const checkboxes = elements.filter((e) => e.type === "checkbox").length;
  const tables = elements.filter((e) => e.type === "table").length;
  const quotes = elements.filter((e) => e.type === "blockquote").length;
  console.log(`Headings: ${headings}, Lists: ${lists}, Checkboxes: ${checkboxes}, Tables: ${tables}, Quotes: ${quotes}\n`);

  return elements;
}

// 인라인 포맷팅 파싱 (볼드, 이탤릭, 코드, 링크, 이미지)
function parseInlineFormatting(text) {
  // HTML 마크업 처리 (특수 마커로 임시 치환)
  text = processHtmlMarkup(text);

  const runs = [];
  let i = 0;

  while (i < text.length) {
    // 개행 마커 처리 (\u0001BR\u0001)
    if (i < text.length - 5 && text.substring(i, i + 6) === "\u0001BR\u0001") {
      runs.push(
        new TextRun({
          text: "\n",
          font: fontFamily,
        })
      );
      i += 6;
      continue;
    }

    // 구분선 마커 처리 (\u0001HR\u0001)
    if (i < text.length - 5 && text.substring(i, i + 6) === "\u0001HR\u0001") {
      runs.push(
        new TextRun({
          text: "─────────────────────",
          font: fontFamily,
        })
      );
      i += 6;
      continue;
    }

    // ![alt](image) 처리 (이미지 링크)
    if (i < text.length - 3 && text[i] === "!" && text[i + 1] === "[") {
      const altEnd = text.indexOf("]", i + 2);
      const urlStart = text.indexOf("(", altEnd);
      const urlEnd = text.indexOf(")", urlStart);

      if (altEnd !== -1 && urlStart !== -1 && urlEnd !== -1) {
        const altText = text.substring(i + 2, altEnd);
        const imageUrl = text.substring(urlStart + 1, urlEnd);

        runs.push(
          new TextRun({
            text: `[이미지: ${altText || imageUrl}]`,
            italics: true,
            color: "666666",
            font: fontFamily,
          })
        );
        i = urlEnd + 1;
        continue;
      }
    }

    // [text](url) 처리 (링크)
    if (i < text.length - 3 && text[i] === "[") {
      const textEnd = text.indexOf("]", i + 1);
      const urlStart = text.indexOf("(", textEnd);
      const urlEnd = text.indexOf(")", urlStart);

      if (textEnd !== -1 && urlStart !== -1 && urlEnd !== -1) {
        const linkText = text.substring(i + 1, textEnd);
        const linkUrl = text.substring(urlStart + 1, urlEnd);

        runs.push(
          new TextRun({
            text: linkText,
            color: "0563C1",
            underline: {
              type: "single",
            },
            font: fontFamily,
          })
        );
        runs.push(
          new TextRun({
            text: ` (${linkUrl})`,
            color: "70AD47",
            font: fontFamily,
          })
        );

        i = urlEnd + 1;
        continue;
      }
    }

    // **bold** 처리
    if (i < text.length - 3 && text.substring(i, i + 2) === "**") {
      const end = text.indexOf("**", i + 2);
      if (end !== -1) {
        const boldText = text.substring(i + 2, end);
        runs.push(
          new TextRun({
            text: boldText,
            bold: true,
            font: fontFamily,
          })
        );
        i = end + 2;
        continue;
      }
    }

    // `code` 처리
    if (text[i] === "`") {
      const end = text.indexOf("`", i + 1);
      if (end !== -1) {
        const codeText = text.substring(i + 1, end);
        runs.push(
          new TextRun({
            text: codeText,
            font: "Consolas",
            size: 21,
            color: "CF222E",
            shading: { fill: "F6F8FA" },
          })
        );
        i = end + 1;
        continue;
      }
    }

    // *italic* 처리 (** 아닌 경우)
    if (text[i] === "*" && (i + 1 >= text.length || text[i + 1] !== "*")) {
      const end = text.indexOf("*", i + 1);
      if (end !== -1) {
        const italicText = text.substring(i + 1, end);
        runs.push(
          new TextRun({
            text: italicText,
            italics: true,
            font: fontFamily,
          })
        );
        i = end + 1;
        continue;
      }
    }

    // 일반 텍스트
    let nextSpecial = text.length;
    for (let j = i + 1; j < text.length; j++) {
      if (
        text[j] === "*" ||
        text[j] === "`" ||
        text[j] === "_" ||
        (text[j] === "[" && text[j - 1] !== "!") ||
        (text[j] === "!" && text[j + 1] === "[")
      ) {
        nextSpecial = j;
        break;
      }
    }
    runs.push(
      new TextRun({
        text: text.substring(i, nextSpecial),
        font: fontFamily,
      })
    );
    i = nextSpecial;
  }

  return runs.length > 0 ? runs : [new TextRun({ text, font: fontFamily })];
}

// DOCX 문서 생성
const elements = parseMarkdown(markdownContent);
const children = [];

// 번호 매김 설정
const numberingConfig = {
  config: [
    {
      reference: "bullet-list",
      levels: [
        {
          level: 0,
          format: LevelFormat.BULLET,
          text: "•",
          alignment: AlignmentType.LEFT,
          style: {
            paragraph: {
              indent: { left: 720, hanging: 360 },
            },
          },
        },
        {
          level: 1,
          format: LevelFormat.BULLET,
          text: "◦",
          alignment: AlignmentType.LEFT,
          style: {
            paragraph: {
              indent: { left: 1080, hanging: 360 },
            },
          },
        },
      ],
    },
  ],
};

// 요소를 DOCX 단락으로 변환
elements.forEach((element) => {
  if (element.type === "heading") {
    const headingLevels = {
      1: HeadingLevel.HEADING_1,
      2: HeadingLevel.HEADING_2,
      3: HeadingLevel.HEADING_3,
      4: HeadingLevel.HEADING_4,
      5: HeadingLevel.HEADING_5,
      6: HeadingLevel.HEADING_6,
    };

    children.push(
      new Paragraph({
        heading: headingLevels[element.level] || HeadingLevel.HEADING_1,
        children: parseInlineFormatting(element.text),
      })
    );
  } else if (element.type === "table") {
    // 테이블 처리
    const headerCells = element.headers.map(
      (header) =>
        new TableCell({
          children: [new Paragraph(parseInlineFormatting(header))],
          shading: { fill: "D3D3D3" },
          verticalAlign: VerticalAlign.CENTER,
        })
    );

    const bodyRows = element.rows.map(
      (row) =>
        new TableRow({
          children: row.map(
            (cell) =>
              new TableCell({
                children: [new Paragraph(parseInlineFormatting(cell))],
              })
          ),
        })
    );

    children.push(
      new Table({
        width: { size: 100, type: "pct" },
        rows: [
          new TableRow({
            children: headerCells,
          }),
          ...bodyRows,
        ],
        borders: {
          top: { style: BorderStyle.SINGLE, size: 1, color: "000000" },
          bottom: { style: BorderStyle.SINGLE, size: 1, color: "000000" },
          left: { style: BorderStyle.SINGLE, size: 1, color: "000000" },
          right: { style: BorderStyle.SINGLE, size: 1, color: "000000" },
          insideHorizontal: {
            style: BorderStyle.SINGLE,
            size: 1,
            color: "000000",
          },
          insideVertical: {
            style: BorderStyle.SINGLE,
            size: 1,
            color: "000000",
          },
        },
      })
    );

    children.push(
      new Paragraph({
        text: "",
        spacing: { after: 200 },
      })
    );
  } else if (element.type === "blockquote") {
    children.push(
      new Paragraph({
        children: parseInlineFormatting(element.text),
        indent: { left: 360 },
        spacing: { before: 200, after: 200 },
        shading: {
          fill: "F6F8FA",
        },
        border: {
          left: {
            color: "0969DA",
            space: 4,
            value: "single",
            size: 24,
          },
        },
      })
    );
  } else if (element.type === "checkbox") {
    const checkmark = element.checked ? "☑" : "☐";
    children.push(
      new Paragraph({
        children: [
          new TextRun({
            text: checkmark + " ",
            font: fontFamily,
            size: 22,
          }),
          ...parseInlineFormatting(element.text),
        ],
        numbering: {
          reference: "bullet-list",
          level: element.level,
        },
        spacing: { before: 40, after: 80 },
      })
    );
  } else if (element.type === "list") {
    children.push(
      new Paragraph({
        children: parseInlineFormatting(element.text),
        numbering: {
          reference: "bullet-list",
          level: element.level,
        },
        spacing: { before: 40, after: 80 },
      })
    );
  } else if (element.type === "paragraph") {
    // 개행 처리 - \n을 기준으로 나누기
    const paragraphLines = element.text.split("\n");
    paragraphLines.forEach((line, index) => {
      if (line.trim() || index > 0) {
        // 빈 줄도 유지
        children.push(
          new Paragraph({
            children: parseInlineFormatting(line),
            spacing: { after: 200 },
          })
        );
      }
    });
  }
});

// 문서 생성
const doc = new Document({
  styles: {
    default: {
      document: {
        run: {
          font: fontFamily,
          size: 22,
        },
        paragraph: {
          spacing: { line: 276, lineRule: "auto" },
        },
      },
    },
    paragraphStyles: [
      {
        id: "Heading1",
        name: "Heading 1",
        basedOn: "Normal",
        next: "Normal",
        quickFormat: true,
        run: {
          font: fontFamily,
          size: 56,
          bold: true,
          color: "000000",
        },
        paragraph: {
          spacing: { before: 400, after: 320 },
        },
      },
      {
        id: "Heading2",
        name: "Heading 2",
        basedOn: "Normal",
        next: "Normal",
        quickFormat: true,
        run: {
          font: fontFamily,
          size: 42,
          bold: true,
          color: "000000",
        },
        paragraph: {
          spacing: { before: 380, after: 200 },
        },
      },
      {
        id: "Heading3",
        name: "Heading 3",
        basedOn: "Normal",
        next: "Normal",
        quickFormat: true,
        run: {
          font: fontFamily,
          size: 32,
          bold: true,
          color: "000000",
        },
        paragraph: {
          spacing: { before: 360, after: 160 },
        },
      },
      {
        id: "Heading4",
        name: "Heading 4",
        basedOn: "Normal",
        next: "Normal",
        quickFormat: true,
        run: {
          font: fontFamily,
          size: 28,
          bold: true,
          color: "000000",
        },
        paragraph: {
          spacing: { before: 320, after: 120 },
        },
      },
    ],
  },
  numbering: numberingConfig,
  sections: [
    {
      properties: {
        page: {
          margin: {
            top: 1440,
            right: 1440,
            bottom: 1440,
            left: 1440,
          },
        },
      },
      children: children,
    },
  ],
});

// 파일 저장
Packer.toBuffer(doc)
  .then((buffer) => {
    try {
      fs.writeFileSync(outputFile, buffer);
      console.log(`✓ DOCX 파일 생성 완료: ${outputFile}`);
      console.log(`\n📊 통계:`);
      console.log(`  총 요소: ${elements.length}개`);
      console.log(`  글꼴: ${fontFamily}`);
      console.log(`\n🔄 사용법: node markdown_to_docx.js [options]`);
      console.log(`  --font "폰트명"  : 사용할 폰트 (기본값: 맑은 고딕)`);
      console.log(`  --input "파일명" : 입력 마크다운 파일 (기본값: input.md)`);
      console.log(`\n📝 예시:`);
      console.log(`  node markdown_to_docx.js`);
      console.log(`  node markdown_to_docx.js --font "나눔고딕"`);
      console.log(`  node markdown_to_docx.js --input "마크다운파일.md"`);
      console.log(`  node markdown_to_docx.js --input "마크다운파일.md" --font "Arial"\n`);
    } catch (error) {
      console.error("❌ 파일 저장 오류:");
      console.error(`  ${error.message}\n`);
      process.exit(1);
    }
  })
  .catch((error) => {
    console.error("❌ DOCX 생성 오류:");
    console.error(`  ${error.message}\n`);
    process.exit(1);
  });
