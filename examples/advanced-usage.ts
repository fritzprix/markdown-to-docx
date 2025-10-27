// examples/advanced-usage.ts
// 라이브러리 고급 사용 예제 (TypeScript)

import * as fs from "fs";
import { parseMarkdown, DocxBuilder } from "../src/index";

async function main() {
  try {
    // 마크다운 파일 읽기
    const content = fs.readFileSync("test_features.md", "utf-8");

    // 마크다운 파싱
    const elements = parseMarkdown(content, { verbose: true });

    // 커스텀 폰트로 DOCX 빌더 생성
    const builder = new DocxBuilder("Arial");

    // 요소 추가
    elements.forEach((element) => {
      builder.addElement(element);
    });

    // 파일 저장
    const file = await builder.save("advanced_output.docx");
    console.log(`✓ 파일 생성됨: ${file}`);
  } catch (error) {
    console.error("오류:", error);
  }
}

main();
