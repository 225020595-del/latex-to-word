import { Document, Packer, Paragraph, TextRun, HeadingLevel, XmlComponent, Math } from "docx";
import { unified } from "unified";
import remarkParse from "remark-parse";
import remarkMath from "remark-math";
import { mathjax } from "mathjax-full/js/mathjax.js";
import { TeX } from "mathjax-full/js/input/tex.js";
import { OmmlVisitor } from "./omml-visitor";

// Initialize MathJax for OMML generation
const tex = new TeX({ packages: ["base", "ams"] });
// We don't need a full document, just input jax. 
// But mathjax-full API requires a document or similar context to parse.
// We can use a minimal setup.
// Actually, `tex.parse` returns MmlNode which is what we need.
// BUT `tex.parse` is internal. 
// The standard way is `mathjax.document(...).convert(...)`.

// Create a visitor instance
const ommlVisitor = new OmmlVisitor();

function convertLatexToOmml(latex: string, display: boolean): any {
  try {
    // 1. Parse LaTeX to MathJax Internal MmlNode
    // We create a new MathDocument for each conversion to ensure clean state or reuse?
    // Reusing is better.
    // Note: `mathjax.document` requires an adaptor.
    // We can use `liteAdaptor`.
    const { liteAdaptor } = require("mathjax-full/js/adaptors/liteAdaptor.js");
    const { RegisterHTMLHandler } = require("mathjax-full/js/handlers/html.js");
    const adaptor = liteAdaptor();
    RegisterHTMLHandler(adaptor);
    
    const html = mathjax.document("", { InputJax: tex });
    
    // Convert to MathItem (which holds the MmlNode tree)
    const mathItem = html.convert(latex, { display: display });
    
    // 2. Visit the MmlNode tree to generate OMML XML string
    const ommlString = ommlVisitor.visitTree(mathItem);
    
    // 3. Return as XmlComponent for docx
    return new XmlComponent(ommlString);
  } catch (e) {
    console.error("OMML conversion error:", e);
    // Fallback to text if conversion fails
    return new TextRun(`[Error: ${latex}]`);
  }
}

// Recursive helper to process AST nodes into Docx elements
function processNode(node: any, parentStyle?: { bold?: boolean; italics?: boolean }): (TextRun | XmlComponent)[] {
  const results: (TextRun | XmlComponent)[] = [];
  
  if (node.type === "text") {
    results.push(new TextRun({ 
      text: node.value, 
      bold: parentStyle?.bold, 
      italics: parentStyle?.italics 
    }));
  } else if (node.type === "inlineMath") {
    // Use OMML converter
    results.push(convertLatexToOmml(node.value, false));
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
      // Block Math -> Display Mode
      // We wrap it in a paragraph
      children.push(new Paragraph({
        children: [
          convertLatexToOmml(node.value, true)
        ],
        alignment: "center", 
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
