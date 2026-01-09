import { Document, Packer, Paragraph, TextRun, ImageRun, HeadingLevel } from "docx";
import { unified } from "unified";
import remarkParse from "remark-parse";
import remarkMath from "remark-math";
import { mathjax } from "mathjax-full/js/mathjax.js";
import { TeX } from "mathjax-full/js/input/tex.js";
import { SVG } from "mathjax-full/js/output/svg.js";
import { liteAdaptor } from "mathjax-full/js/adaptors/liteAdaptor.js";
import { RegisterHTMLHandler } from "mathjax-full/js/handlers/html.js";

// Initialize MathJax
const adaptor = liteAdaptor();
RegisterHTMLHandler(adaptor);

const tex = new TeX({ packages: ["base", "ams"] });
const svg = new SVG({ fontCache: "none" });
const html = mathjax.document("", { InputJax: tex, OutputJax: svg });

// 1x1 transparent PNG fallback for older Word versions
const FALLBACK_IMAGE_BASE64 = "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNk+M9QDwADhgGAWjR9awAAAABJRU5ErkJggg==";
const FALLBACK_BUFFER = Buffer.from(FALLBACK_IMAGE_BASE64, "base64");

// Helper to convert LaTeX to SVG string and extract dimensions
function convertLatexToSvg(latex: string, isInline: boolean): { svg: string; width: number; height: number } | null {
  try {
    const node = html.convert(latex, {
      display: !isInline,
      em: 16,
      ex: 8,
      containerWidth: 80 * 16,
    });
    
    const fullHtml = adaptor.innerHTML(node);
    const svgMatch = fullHtml.match(/<svg[\s\S]*?<\/svg>/);
    
    if (!svgMatch) {
      console.warn("No SVG found in MathJax output:", fullHtml);
      return null;
    }
    
    const svgString = svgMatch[0];
    const widthMatch = svgString.match(/width="([\d.]+)ex"/);
    const heightMatch = svgString.match(/height="([\d.]+)ex"/);
    
    let width = 100;
    let height = 30;
    
    const EX_TO_PX = 10; 
    
    if (widthMatch && widthMatch[1]) {
      width = parseFloat(widthMatch[1]) * EX_TO_PX;
    }
    if (heightMatch && heightMatch[1]) {
      height = parseFloat(heightMatch[1]) * EX_TO_PX;
    }
    
    width = Math.max(width, 10);
    height = Math.max(height, 10);

    return { svg: svgString, width, height };
  } catch (e) {
    console.error("MathJax conversion error:", e);
    return null;
  }
}

// Recursive helper to process AST nodes into Docx elements
function processNode(node: any, parentStyle?: { bold?: boolean; italics?: boolean }): (TextRun | ImageRun)[] {
  const results: (TextRun | ImageRun)[] = [];
  
  // Handle current node content
  if (node.type === "text") {
    results.push(new TextRun({ 
      text: node.value, 
      bold: parentStyle?.bold, 
      italics: parentStyle?.italics 
    }));
  } else if (node.type === "inlineMath") {
    const result = convertLatexToSvg(node.value, true);
    if (result) {
      results.push(new ImageRun({
        data: Buffer.from(result.svg),
        transformation: { width: result.width, height: result.height },
        type: "svg",
        fallback: { data: FALLBACK_BUFFER, type: "png" }
      }));
    } else {
      results.push(new TextRun({ text: `$${node.value}$`, color: "red" }));
    }
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
    // Fallback for other inline containers
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
        // Recursive call for nested lists
        paragraphs.push(...processListItems(child, level + 1));
      } else {
        // Handle other block types inside list items if needed
        // For now, try to process as inline content if possible
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

  // Traverse top-level block nodes
  for (const node of (ast as any).children) {
    if (node.type === "paragraph") {
      const runs = processNode(node);
      children.push(new Paragraph({ children: runs }));
    } else if (node.type === "math") {
      // Block Math
      const result = convertLatexToSvg(node.value, false);
      if (result) {
        children.push(new Paragraph({
          children: [new ImageRun({
            data: Buffer.from(result.svg),
            transformation: { width: result.width, height: result.height },
            type: "svg",
            fallback: { data: FALLBACK_BUFFER, type: "png" }
          })],
          alignment: "center",
          spacing: { before: 200, after: 200 },
        }));
      } else {
         children.push(new Paragraph({ children: [new TextRun({ text: `$$${node.value}$$`, color: "red" })] }));
      }
    } else if (node.type === "heading") {
      const level = node.depth as number;
      // Map depth to HeadingLevel enum
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
      // Improved List Handling: Support nested lists
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
