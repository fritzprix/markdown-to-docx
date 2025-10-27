# AI Coding Agent Instructions for markdown-to-docx

## Architecture Overview

**markdown-to-docx** is a TypeScript-based library and CLI that converts Markdown files to Word (DOCX) documents while preserving structure and formatting.

### Data Flow Pipeline

1. **Input** → `markdownParser.ts`: Parse markdown text into typed element objects
2. **Process** → `inlineFormatter.ts`: Convert markdown inline syntax (**bold**, *italic*, `code`, links) to DOCX TextRun objects
3. **Build** → `docxBuilder.ts`: Accumulate elements and generate Word document via `docx` library
4. **Output** → `index.ts` or `cli.ts`: Save DOCX file or return stream

### Key Files & Responsibilities

- **`src/markdownParser.ts`**: Core parser with regex-based element detection (headings, lists, tables, checkboxes, images). **Special handling**: HTML markup (tags, comments) temporarily converted to `\u0001` markers to prevent interference with parsing.
- **`src/inlineFormatter.ts`**: Text styling layer that recurses through markdown inline syntax to build TextRun arrays compatible with `docx` library.
- **`src/docxBuilder.ts`**: Stateful builder class that accumulates elements and generates styled Word documents. Uses `docx` library's Document/Paragraph/Table APIs. Manages image loading and metadata.
- **`src/imageLoader.ts`**: Image loading utility (NEW). Handles local file I/O, HTTP(S) remote image downloads, MIME type detection, and error handling with graceful fallback.
- **`src/index.ts`**: Public API - single entry point for library usage. Orchestrates parser → builder → file save flow. Passes markdown file directory to DocxBuilder for relative image path resolution.
- **`src/cli.ts`**: CLI entry point. Parses `--input`, `--output`, `--font` arguments and delegates to `index.ts`.

## TypeScript Patterns & Conventions

### Type Safety via Union Types
Elements use discriminated union pattern:
```typescript
export type ParsedElement = 
  | TableElement | CheckboxElement | HeadingElement 
  | ListElement | BlockquoteElement | ParagraphElement | ImageElement;
```
When adding new element types, update `ParsedElement` union and add corresponding `add<Type>` method to DocxBuilder.

### Interface-First Design
All element types are strict interfaces (not loose `any` usage). Critical types in `markdownParser.ts` export:
```typescript
export interface HeadingElement { type: "heading"; level: number; text: string; }
```

### ESLint Disable Comments
Use ESLint disable where typing limitation exists (currently in `docxBuilder.ts` for HeadingLevel map). Keep comments minimal and specific.

## Developer Workflows

### Build & Test Cycle
```bash
npm run build          # Compile TypeScript → dist/
npm run lint           # Check code style (must pass before commit)
npm run format:check   # Verify Prettier formatting
npm run dev -- --input tests/test_features.md  # Quick test with ts-node
```

### Before Committing
```bash
npm run lint:fix       # Auto-fix linting issues
npm run format         # Auto-format code
npm run build          # Verify compilation
```

### Adding Features
1. Update parser in `markdownParser.ts` with new element interface
2. Add formatter logic in `inlineFormatter.ts` if inline styling needed
3. Add builder method in `docxBuilder.ts` (e.g., `addCustomElement()`)
4. Route in `docxBuilder.addElement()` dispatcher
5. Test with `tests/test_*.md` files and verify output DOCX visually

## Critical Implementation Details

### HTML Markup Handling (Fragile)
Markdown parsers easily break when encountering HTML tags. Solution: temporarily replace `<br/>`, `<hr/>`, etc. with control characters (`\u0001BR\u0001`) **before** parsing, restore afterward via `restoreHtmlMarkup()`. See `src/markdownParser.ts` lines 67-75. **Never remove this layer** without refactoring the parser.

### Windows Line Endings
Test files use CRLF; parser normalizes via `.replace(/\r\n/g, '\n')` at start of `parseMarkdown()`. Essential for cross-platform consistency.

### Font Consistency
All formatting functions accept `fontFamily` parameter (default: "맑은 고딕"). Pass through every call chain: CLI → index → builder → formatters. Inconsistent fonts break document appearance.

### DOCX Library Limitations
- Images are loaded but rendered as `[Image: alt_text]` placeholders in current implementation.
- For actual embedded images in DOCX: Requires `docx` library's `Media` API (complex setup) or alternatives like `PptxGenJs`.
- HTTP(S) image downloads supported via `imageLoader.ts` with timeouts and MIME type detection.
- Local images loaded via `basePath` parameter in DocxBuilder (relative to markdown file directory).

## Image Processing

### Implementation Flow
1. **Parsing**: `parseImageMarkdown()` in `markdownParser.ts` detects `![alt](src)` syntax (standalone lines only, not inline).
2. **Loading**: `loadImage()` in `imageLoader.ts` determines if `src` is local path or HTTP(S) URL.
   - **Local**: Read file sync with `fs.readFileSync()`, resolve relative paths via `basePath`.
   - **Remote**: Async fetch via `fetch()` (Node.js 18+) or http/https modules (Node.js <18) with 10s timeout.
3. **Building**: `addImage()` in `docxBuilder.ts` currently renders image as placeholder text (future enhancement for actual embedding).
4. **Error handling**: Failed image loads gracefully display `[Image Error: reason]` instead of crashing.

### Adding Real Image Support (Future)
To embed actual images in DOCX:
1. Store image buffers in `DocxBuilder.media: Map<string, Buffer>`.
2. Use `docx` library's undocumented `_imageIds` or switch to document-level `Media` registration.
3. Insert `ImageRun` objects in paragraph children (requires `docx` v10+).
4. Update `build()` method to include media section in document XML.
5. Test with compressed/optimized images to manage file size.

### Image Syntax & Options
Supported markdown syntax for images:
- Local file paths: relative or absolute paths to image files
- Remote HTTP URLs: full URL to image hosted online
- Optional sizing: dimensions specified in curly braces

Common patterns (see tests/test_features.md for examples):
1. Standalone image: image link on its own line
2. Local paths resolved relative to markdown file directory
3. Remote images auto-downloaded with 10s timeout
4. Fallback to text placeholder on load failure

## Code Quality Tooling

- **ESLint**: `eslint.config.cjs` (v9+ format). Rules: TypeScript recommended + warnings for `any` type.
- **Prettier**: `.prettierrc` (100 char line width, 2-space indent, trailing commas).
- **TypeScript**: `tsconfig.json` (strict mode enabled, target: ES2020).

Run formatters before every commit; CI/CD will enforce on PR.

## Testing Strategy

No automated test framework. Manual testing via:
- `npm run dev -- --input tests/test_features.md` → opens generated DOCX visually
- Test files in `tests/` directory cover: tables, links, images, checkboxes, mixed formatting, quotes.
- Verify DOCX visually in Word to catch subtle styling issues.

## Distribution & npm Packaging

- **package.json**: `"main": "dist/index.js"`, `"bin": { "markdown-to-docx": "dist/cli.js" }`
- **Publish**: `npm publish` (prepublishOnly script auto-builds). Users install via `npm install -g markdown-to-docx`.
- **License**: BSD-3-Clause (see LICENSE file).

## Common Pitfalls

1. **Forgetting TypeScript compilation**: CLI runs `dist/` files, not `src/`. Always `npm run build` after changes.
2. **Font parameter dropped**: Ensure `fontFamily` flows through all function calls in DocxBuilder methods.
3. **Parser edge cases**: Markdown is ambiguous (e.g., `>` could be blockquote or HTML). Test new parsing logic extensively.
4. **Breaking HTML marker layer**: The `\u0001` replacement system is fragile; refactor carefully if modifying parser.

## Resources

- **docx library**: https://www.npmjs.com/package/docx - Low-level Word document API
- **Markdown spec**: https://www.markdownguide.org/ - Reference for supported syntax
- **TypeScript**: https://www.typescriptlang.org/ - Strict mode best practices
- **DEVELOPMENT.md**: Full dev guide with npm script reference and npm publish steps
