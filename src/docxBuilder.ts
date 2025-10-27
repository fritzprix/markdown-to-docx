// src/docxBuilder.ts
// DOCX 문서 생성 로직

import {
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
} from "docx";
import * as fs from "fs";
import { parseInlineFormatting } from "./inlineFormatter";
import type {
  ParsedElement,
  HeadingElement,
  TableElement,
  BlockquoteElement,
  CheckboxElement,
  ListElement,
  ParagraphElement,
} from "./markdownParser";

interface NumberingConfig {
  config: Array<{
    reference: string;
    levels: Array<{
      level: number;
      format: typeof LevelFormat.BULLET;
      text: string;
      alignment: typeof AlignmentType.LEFT;
      style: {
        paragraph: {
          indent: { left: number; hanging: number };
        };
      };
    }>;
  }>;
}

export class DocxBuilder {
  private fontFamily: string;
  private children: (Paragraph | Table)[] = [];
  private numberingConfig: NumberingConfig;

  constructor(fontFamily: string = "맑은 고딕") {
    this.fontFamily = fontFamily;
    this.numberingConfig = {
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
  }

  addElement(element: ParsedElement): void {
    if (element.type === "heading") {
      this.addHeading(element as HeadingElement);
    } else if (element.type === "table") {
      this.addTable(element as TableElement);
    } else if (element.type === "blockquote") {
      this.addBlockquote(element as BlockquoteElement);
    } else if (element.type === "checkbox") {
      this.addCheckbox(element as CheckboxElement);
    } else if (element.type === "list") {
      this.addList(element as ListElement);
    } else if (element.type === "paragraph") {
      this.addParagraph(element as ParagraphElement);
    }
  }

  private addHeading(element: HeadingElement): void {
    const headingLevelMap: Record<
      number,
      | typeof HeadingLevel.HEADING_1
      | typeof HeadingLevel.HEADING_2
      | typeof HeadingLevel.HEADING_3
      | typeof HeadingLevel.HEADING_4
      | typeof HeadingLevel.HEADING_5
      | typeof HeadingLevel.HEADING_6
    > = {
      1: HeadingLevel.HEADING_1,
      2: HeadingLevel.HEADING_2,
      3: HeadingLevel.HEADING_3,
      4: HeadingLevel.HEADING_4,
      5: HeadingLevel.HEADING_5,
      6: HeadingLevel.HEADING_6,
    };

    const headingLevel = headingLevelMap[element.level] ?? HeadingLevel.HEADING_1;

    this.children.push(
      new Paragraph({
        heading: headingLevel,
        children: parseInlineFormatting(element.text, this.fontFamily),
      })
    );
  }

  private addTable(element: TableElement): void {
    const headerCells = element.headers.map(
      (header) =>
        new TableCell({
          children: [
            new Paragraph({
              children: parseInlineFormatting(header, this.fontFamily),
            }),
          ],
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
                children: [
                  new Paragraph({
                    children: parseInlineFormatting(cell, this.fontFamily),
                  }),
                ],
              })
          ),
        })
    );

    this.children.push(
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

    this.children.push(
      new Paragraph({
        text: "",
        spacing: { after: 200 },
      })
    );
  }

  private addBlockquote(element: BlockquoteElement): void {
    this.children.push(
      new Paragraph({
        children: parseInlineFormatting(element.text, this.fontFamily),
        indent: { left: 360 },
        spacing: { before: 200, after: 200 },
        shading: {
          fill: "F6F8FA",
        },
        border: {
          left: {
            color: "0969DA",
            space: 4,
            style: BorderStyle.SINGLE,
            size: 24,
          },
        },
      })
    );
  }

  private addCheckbox(element: CheckboxElement): void {
    const checkmark = element.checked ? "☑" : "☐";
    this.children.push(
      new Paragraph({
        children: [
          new TextRun({
            text: checkmark + " ",
            font: this.fontFamily,
            size: 22,
          }),
          ...parseInlineFormatting(element.text, this.fontFamily),
        ],
        numbering: {
          reference: "bullet-list",
          level: element.level,
        },
        spacing: { before: 40, after: 80 },
      })
    );
  }

  private addList(element: ListElement): void {
    this.children.push(
      new Paragraph({
        children: parseInlineFormatting(element.text, this.fontFamily),
        numbering: {
          reference: "bullet-list",
          level: element.level,
        },
        spacing: { before: 40, after: 80 },
      })
    );
  }

  private addParagraph(element: ParagraphElement): void {
    const paragraphLines = element.text.split("\n");
    paragraphLines.forEach((line, index) => {
      if (line.trim() || index > 0) {
        this.children.push(
          new Paragraph({
            children: parseInlineFormatting(line, this.fontFamily),
            spacing: { after: 200 },
          })
        );
      }
    });
  }

  build(): Document {
    return new Document({
      styles: {
        default: {
          document: {
            run: {
              font: this.fontFamily,
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
              font: this.fontFamily,
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
              font: this.fontFamily,
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
              font: this.fontFamily,
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
              font: this.fontFamily,
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
      numbering: this.numberingConfig,
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
          children: this.children,
        },
      ],
    });
  }

  async save(outputFile: string): Promise<string> {
    const doc = this.build();
    const buffer = await Packer.toBuffer(doc);
    fs.writeFileSync(outputFile, buffer);
    return outputFile;
  }
}
