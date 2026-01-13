import { Document, Packer, Paragraph, TextRun, HeadingLevel, XmlComponent, Math, MathRun } from "docx";
import { unified } from "unified";
import remarkParse from "remark-parse";
import remarkMath from "remark-math";
import { latexToWordLinear } from "./latex-transformer";

// Revert to Linear Format for stability, as raw XML injection is risky in browser/node hybrid envs
// and caused import issues.
// We will rely on the improved `latex-transformer` to make the linear format robust.

// Recursive helper to process AST nodes into Docx elements
function processNode(node: any, parentStyle?: { bold?: boolean; italics?: boolean }): (TextRun | Math)[] {
  const results: (TextRun | Math)[] = [];
  
  if (node.type === "text") {
    results.push(new TextRun({ 
      text: node.value, 
      bold: parentStyle?.bold, 
      italics: parentStyle?.italics 
    }));
  } else if (node.type === "inlineMath") {
    // Improved Linear Format usage
    results.push(new Math({
      children: [
        new MathRun(latexToWordLinear(node.value))
      ]
    }));
  } else if (node.type === "emphasis" || node.type === "strong") {
    const isBold = node.type === "strong" || parentStyle?.bold;
    const isItalic = node.type === "emphasis" || parentStyle?.italics;
    
    if (node.children) {
      for (const child of node.children) {
        results.push(...processNode(child, { bold: isBold, italics: isItalic }));
      }
    }
  } else if (node.type === "link") {
      if (node.children) {
        for (const child of node.children) {
           results.push(...processNode(child, { ...parentStyle }));
        }
      }
  } else if (node.children) {
    for (const child of node.children) {
      results.push(...processNode(child, parentStyle));
    }
  }

  return results;
}

// New helper function to process list items recursively
function processListItems(node: any, level: number = 0): Paragraph[] {
  const paragraphs: Paragraph[] = [];
  
  node.children.forEach((listItem: any) => {
    listItem.children.forEach((child: any) => {
      if (child.type === "paragraph" || child.type === "text") {
        const runs = processNode(child);
        paragraphs.push(new Paragraph({
          children: runs,
          bullet: {
            level: level 
          }
        }));
      } else if (child.type === "list") {
        paragraphs.push(...processListItems(child, level + 1));
      } else {
        const runs = processNode(child);
        if (runs.length > 0) {
           paragraphs.push(new Paragraph({
             children: runs,
             bullet: {
               level: level
             }
           }));
        }
      }
    });
  });
  
  return paragraphs;
}

export async function createDocxFromMarkdown(markdown: string): Promise<string> {
  const processor = unified().use(remarkParse).use(remarkMath);
  const ast = processor.parse(markdown);

  const children: Paragraph[] = [];

  for (const node of (ast as any).children) {
    if (node.type === "paragraph") {
      const runs = processNode(node);
      children.push(new Paragraph({ children: runs }));
    } else if (node.type === "math") {
      // Block Math
      children.push(new Paragraph({
        children: [
          new Math({
            children: [
              new MathRun(latexToWordLinear(node.value))
            ]
          })
        ],
        alignment: "center", // Center block math
      }));
    } else if (node.type === "heading") {
      const level = node.depth as number;
      let headingLevel: any = HeadingLevel.HEADING_1;
      if (level === 2) headingLevel = HeadingLevel.HEADING_2;
      if (level === 3) headingLevel = HeadingLevel.HEADING_3;
      if (level === 4) headingLevel = HeadingLevel.HEADING_4;
      if (level === 5) headingLevel = HeadingLevel.HEADING_5;
      if (level === 6) headingLevel = HeadingLevel.HEADING_6;
      
      const runs = processNode(node);
      children.push(new Paragraph({
        children: runs,
        heading: headingLevel,
        spacing: { before: 240, after: 120 }
      }));
    } else if (node.type === "list") {
      children.push(...processListItems(node, 0));
    }
  }

  const doc = new Document({
    sections: [{
      properties: {},
      children: children,
    }],
  });

  const buffer = await Packer.toBuffer(doc);
  return buffer.toString("base64");
}
