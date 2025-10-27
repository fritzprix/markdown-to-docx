// src/inlineFormatter.ts
// 인라인 포맷팅 (bold, italic, code, links, images)

import { TextRun } from "docx";
import { processHtmlMarkup } from "./markdownParser";

export function parseInlineFormatting(text: string, fontFamily: string = "맑은 고딕"): TextRun[] {
  // HTML 마크업 처리 (특수 마커로 임시 치환)
  text = processHtmlMarkup(text);

  const runs: TextRun[] = [];
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
