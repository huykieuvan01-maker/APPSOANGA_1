import { 
  Document, Paragraph, TextRun, Table, TableRow, TableCell, 
  AlignmentType, HeadingLevel, LevelFormat, WidthType, 
  BorderStyle, ShadingType, Packer 
} from 'docx';

function parseTextRuns(text: string): TextRun[] {
  const parts = text.split('**');
  const runs: TextRun[] = [];
  
  for (let i = 0; i < parts.length; i++) {
    const isBold = i % 2 === 1;
    const partText = parts[i];
    if (!partText) continue;
    
    const italicParts = partText.split('*');
    for (let j = 0; j < italicParts.length; j++) {
      const isItalic = j % 2 === 1;
      const italicText = italicParts[j];
      if (!italicText) continue;
      
      runs.push(new TextRun({
        text: italicText,
        bold: isBold,
        italics: isItalic,
        font: "Arial",
        size: 24 // 12pt
      }));
    }
  }
  
  return runs;
}

export async function exportMarkdownToDocx(markdown: string): Promise<Blob> {
  const lines = markdown.split(/\r?\n/);
  const children: any[] = [];
  
  let inTable = false;
  let tableRows: string[][] = [];
  let tableHeaders: string[] = [];
  let isHeaderRow = true;

  const processTable = () => {
    if (tableRows.length === 0 && tableHeaders.length === 0) return;
    
    const docxRows: TableRow[] = [];
    
    // Add Header Row
    if (tableHeaders.length > 0) {
      const headerCells = tableHeaders.map(cellText => {
        return new TableCell({
          children: [
            new Paragraph({
              children: [
                new TextRun({
                  text: cellText.trim(),
                  bold: true,
                  color: "FFFFFF",
                  font: "Arial",
                  size: 24
                })
              ],
              alignment: AlignmentType.LEFT
            })
          ],
          shading: {
            fill: "1F4E79", // Dark Blue
            type: ShadingType.CLEAR
          },
          margins: { top: 120, bottom: 120, left: 180, right: 180 }
        });
      });
      docxRows.push(new TableRow({ children: headerCells, tableHeader: true }));
    }
    
    // Add Data Rows
    for (const rowData of tableRows) {
      const dataCells = rowData.map((cellText, idx) => {
        return new TableCell({
          children: [
            new Paragraph({
              children: parseTextRuns(cellText.trim())
            })
          ],
          shading: idx % 2 === 1 ? { fill: "F9FBFD" } : undefined, // Alternate zebra pattern
          margins: { top: 100, bottom: 100, left: 180, right: 180 }
        });
      });
      docxRows.push(new TableRow({ children: dataCells }));
    }
    
    const table = new Table({
      rows: docxRows,
      width: {
        size: 100,
        type: WidthType.PERCENTAGE
      },
      borders: {
        top: { style: BorderStyle.SINGLE, size: 6, color: "D0D0D0" },
        bottom: { style: BorderStyle.SINGLE, size: 6, color: "D0D0D0" },
        left: { style: BorderStyle.SINGLE, size: 6, color: "D0D0D0" },
        right: { style: BorderStyle.SINGLE, size: 6, color: "D0D0D0" },
        insideHorizontal: { style: BorderStyle.SINGLE, size: 4, color: "E0E0E0" },
        insideVertical: { style: BorderStyle.SINGLE, size: 4, color: "E0E0E0" }
      }
    });
    
    children.push(table);
    children.push(new Paragraph({ spacing: { after: 120 } }));
    
    // Reset table parser state
    inTable = false;
    tableRows = [];
    tableHeaders = [];
    isHeaderRow = true;
  };

  for (let lineIdx = 0; lineIdx < lines.length; lineIdx++) {
    const line = lines[lineIdx].trim();
    
    // 1. Table Detection
    if (line.startsWith('|')) {
      inTable = true;
      // Skip separator lines like |---|---|
      if (line.includes('---')) {
        isHeaderRow = false;
        continue;
      }
      
      const columns = line.split('|').map(c => c.trim()).filter((_, idx, arr) => idx > 0 && idx < arr.length - 1);
      
      if (isHeaderRow) {
        tableHeaders = columns;
      } else {
        tableRows.push(columns);
      }
      continue;
    } else if (inTable) {
      processTable();
    }
    
    // Skip empty lines
    if (line === '') {
      children.push(new Paragraph({ spacing: { after: 120 } }));
      continue;
    }
    
    // 2. Heading Detection
    if (line.startsWith('# ')) {
      children.push(new Paragraph({
        heading: HeadingLevel.TITLE,
        spacing: { before: 240, after: 120 },
        children: [
          new TextRun({
            text: line.substring(2),
            bold: true,
            color: "000000",
            font: "Arial",
            size: 40 // 20pt
          })
        ]
      }));
    } else if (line.startsWith('## ')) {
      children.push(new Paragraph({
        heading: HeadingLevel.HEADING_1,
        spacing: { before: 240, after: 120 },
        children: [
          new TextRun({
            text: line.substring(3),
            bold: true,
            color: "1F4E79",
            font: "Arial",
            size: 32 // 16pt
          })
        ]
      }));
    } else if (line.startsWith('### ')) {
      children.push(new Paragraph({
        heading: HeadingLevel.HEADING_2,
        spacing: { before: 180, after: 120 },
        children: [
          new TextRun({
            text: line.substring(4),
            bold: true,
            color: "2E75B6",
            font: "Arial",
            size: 28 // 14pt
          })
        ]
      }));
    } else if (line.startsWith('#### ')) {
      children.push(new Paragraph({
        heading: HeadingLevel.HEADING_3,
        spacing: { before: 120, after: 80 },
        children: [
          new TextRun({
            text: line.substring(5),
            bold: true,
            color: "000000",
            font: "Arial",
            size: 24 // 12pt
          })
        ]
      }));
    }
    // 3. Lists
    else if (line.startsWith('- ') || line.startsWith('* ')) {
      children.push(new Paragraph({
        numbering: {
          reference: "bullet-list",
          level: 0
        },
        children: parseTextRuns(line.substring(2))
      }));
    } else if (/^\d+\.\s/.test(line)) {
      const match = line.match(/^(\d+)\.\s(.*)/);
      if (match) {
        children.push(new Paragraph({
          numbering: {
            reference: "number-list",
            level: 0
          },
          children: parseTextRuns(match[2])
        }));
      }
    }
    // 4. Blockquotes
    else if (line.startsWith('> ')) {
      children.push(new Paragraph({
        indent: { left: 720 },
        spacing: { after: 120 },
        children: parseTextRuns(line.substring(2))
      }));
    }
    // 5. Normal paragraphs
    else {
      children.push(new Paragraph({
        spacing: { after: 120 },
        children: parseTextRuns(line)
      }));
    }
  }

  // Ensure last table is processed if document ends with a table
  if (inTable) {
    processTable();
  }

  const doc = new Document({
    numbering: {
      config: [
        {
          reference: "bullet-list",
          levels: [
            {
              level: 0,
              format: LevelFormat.BULLET,
              text: "•",
              alignment: AlignmentType.LEFT,
              style: { paragraph: { indent: { left: 540, hanging: 270 } } }
            }
          ]
        },
        {
          reference: "number-list",
          levels: [
            {
              level: 0,
              format: LevelFormat.DECIMAL,
              text: "%1.",
              alignment: AlignmentType.LEFT,
              style: { paragraph: { indent: { left: 540, hanging: 270 } } }
            }
          ]
        }
      ]
    },
    sections: [
      {
        properties: {
          page: {
            margin: {
              top: 1440, // 1 inch
              bottom: 1440,
              left: 1440,
              right: 1440
            }
          }
        },
        children: children
      }
    ]
  });

  return await Packer.toBlob(doc);
}
