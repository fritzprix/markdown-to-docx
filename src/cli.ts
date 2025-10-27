#!/usr/bin/env node

// src/cli.ts
// 명령줄 인터페이스 진입점

import { convertMarkdownToDOCX } from "./index";

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
  } else if (args[i] === "--output" && args[i + 1]) {
    outputFile = args[i + 1];
  }
}

console.log(`사용 폰트: ${fontFamily}`);
console.log(`입력: ${inputFile}\n`);

// 변환 실행
convertMarkdownToDOCX(inputFile, outputFile, { fontFamily, verbose: true })
  .then(() => {
    console.log(`🔄 사용법: markdown-to-docx [options]`);
    console.log(`  --font "폰트명"    : 사용할 폰트 (기본값: 맑은 고딕)`);
    console.log(`  --input "파일명"   : 입력 마크다운 파일 (기본값: input.md)`);
    console.log(`  --output "파일명"  : 출력 DOCX 파일`);
    console.log(`\n📝 예시:`);
    console.log(`  markdown-to-docx`);
    console.log(`  markdown-to-docx --font "나눔고딕"`);
    console.log(`  markdown-to-docx --input "마크다운파일.md"`);
    console.log(
      `  markdown-to-docx --input "마크다운파일.md" --output "결과.docx" --font "Arial"\n`
    );
  })
  .catch((error: Error) => {
    console.error("\n❌ 오류 발생:");
    console.error(`\n  ${error.message}\n`);
    console.error("💡 해결 방법:");
    console.error(`  1. 파일이 현재 디렉토리에 있는지 확인하세요.`);
    console.error(`     명령어: dir *.md  (또는 ls *.md)\n`);
    console.error(`  2. 파일명을 정확히 지정하세요.`);
    console.error(`     명령어: markdown-to-docx --input "파일명.md"\n`);
    console.error(`  3. 파일이 UTF-8 인코딩인지 확인하세요.\n`);
    console.error("📝 사용법:");
    console.error(`  기본 (input.md 찾기):`);
    console.error(`    markdown-to-docx\n`);
    console.error(`  특정 파일 지정:`);
    console.error(`    markdown-to-docx --input "마크다운파일.md"\n`);
    console.error(`  폰트 변경:`);
    console.error(`    markdown-to-docx --font "나눔고딕"\n`);
    process.exit(1);
  });
