// examples/ts-usage.ts
// TypeScript 라이브러리 사용 예제

import { convertMarkdownToDOCX, DocxBuilder, parseMarkdown } from "../src/index";

async function basicExample() {
  console.log("=== 기본 예제 ===\n");

  // 간단한 변환
  await convertMarkdownToDOCX("test_features.md", "example_ts_output.docx", {
    fontFamily: "맑은 고딕",
    verbose: true,
  });
}

async function advancedExample() {
  console.log("\n=== 고급 예제 (커스텀 빌더) ===\n");

  const markdown = `# TypeScript 예제

이것은 **TypeScript**로 작성된 예제입니다.

- 첫 번째 항목
- 두 번째 항목
  - 중첩된 항목

[링크 예제](https://example.com)
`;

  const builder = new DocxBuilder("Arial");
  const elements = parseMarkdown(markdown, { verbose: false });

  elements.forEach((element) => {
    builder.addElement(element);
  });

  await builder.save("example_ts_custom_output.docx");
  console.log("✓ 커스텀 DOCX 생성 완료: example_ts_custom_output.docx\n");
}

// 실행
(async () => {
  try {
    await basicExample();
    await advancedExample();
  } catch (error) {
    if (error instanceof Error) {
      console.error("오류:", error.message);
    }
  }
})();
