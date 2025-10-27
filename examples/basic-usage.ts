// examples/basic-usage.ts
// 라이브러리 기본 사용 예제 (TypeScript)

import { convertMarkdownToDOCX } from "../src/index";

async function main() {
  // 기본 사용
  const outputFile = await convertMarkdownToDOCX("test_features.md", "example_output.docx", {
    fontFamily: "맑은 고딕",
    verbose: true,
  });

  console.log(`✓ 완료: ${outputFile}`);
}

main().catch(console.error);
