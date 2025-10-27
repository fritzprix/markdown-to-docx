// src/index.ts
// 메인 라이브러리 진입점

import * as fs from "fs";
import { parseMarkdown, type ParseOptions } from "./markdownParser";
import { parseInlineFormatting } from "./inlineFormatter";
import { DocxBuilder } from "./docxBuilder";

interface ConvertOptions extends ParseOptions {
  fontFamily?: string;
}

/**
 * 마크다운 파일을 DOCX로 변환합니다.
 * @param inputFile - 입력 마크다운 파일 경로
 * @param outputFile - 출력 DOCX 파일 경로
 * @param options - 옵션
 * @returns 생성된 파일 경로
 */
export async function convertMarkdownToDOCX(
  inputFile: string,
  outputFile: string,
  options: ConvertOptions = {}
): Promise<string> {
  const fontFamily = options.fontFamily || "맑은 고딕";
  const verbose = options.verbose !== false;

  // 파일 읽기
  if (!fs.existsSync(inputFile)) {
    throw new Error(`파일을 찾을 수 없습니다: ${inputFile}`);
  }

  const content = fs.readFileSync(inputFile, "utf-8");

  // 마크다운 파싱
  const elements = parseMarkdown(content, { verbose });

  // DOCX 생성
  const builder = new DocxBuilder(fontFamily);
  elements.forEach((element) => {
    builder.addElement(element);
  });

  // 파일 저장
  const result = await builder.save(outputFile);

  if (verbose) {
    console.log(`✓ DOCX 파일 생성 완료: ${result}`);
    console.log(`📊 통계:`);
    console.log(`  총 요소: ${elements.length}개`);
    console.log(`  글꼴: ${fontFamily}\n`);
  }

  return result;
}

export { parseMarkdown, parseInlineFormatting, DocxBuilder };
