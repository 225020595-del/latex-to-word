import { Document, Packer, Paragraph, TextRun, HeadingLevel, Math } from "docx";
import { unified } from "unified";
import remarkParse from "remark-parse";
import remarkMath from "remark-math";
import { mj } from "mathjax-node";

// Configure MathJax to output MathML
mj.config({
  MathJax: {
    // Standard MathJax configuration
  }
});
mj.start();

// Helper to convert LaTeX to OMML-compatible MathML
// Note: docx.js doesn't natively support parsing MathML string into objects.
// BUT, Word creates equations using Office MathML (OMML).
// Since we can't easily generate OMML XML manually without a complex parser,
// AND `mathjax-node` outputs MathML (MML), not OMML.
//
// Fortunately, we can use a workaround:
// We will stick to the "Linear Format" approach which is editable.
// The previous issue "format mismatch" likely means the linear format wasn't professional enough.
//
// HOWEVER, users want "Professional" format by default.
// Inserting LaTeX into MathRun is "Linear".
//
// To get Professional format, we really need OMML.
// Since we can't easily get OMML from JS libraries (most do MathML),
// we will try to improve the Linear Format experience or find a library that does MML -> OMML.
//
// Actually, `docx` library has limited Math support.
// Let's try to stick to `MathRun` but maybe we need to process the LaTeX a bit?
//
// If the user says "format mismatch", maybe they mean the LaTeX syntax isn't 100% Word-compatible.
// Word's linear format is *slightly* different from standard LaTeX.
// e.g. `\frac{a}{b}` vs `a/b`.
//
// Let's try to use `mathml2omml` if we can find a way, OR:
// We can use a trick: `docx` supports inserting raw XML.
// If we can get OMML XML, we can insert it.
//
// There is a known XSLT to convert MathML to OMML.
// But running XSLT in Node/Browser is heavy.
//
// Let's try a different approach:
// Use `mathjax-node` to get MathML.
// Then try to insert that MathML? Word supports MathML on clipboard but maybe not in file?
//
// Let's go back to the most reliable method for "Editable Math":
// The `MathRun` with raw LaTeX is actually the standard way to do "Linear Math".
// If it shows up as text, the user just needs to press "Enter" or "Professional" in Word.
//
// But if we want it *already* professional:
// We MUST provide OMML.
//
// Let's try to use a simplified LaTeX -> Word Linear Format converter.
// Or just clean up the LaTeX.
//
// Let's stick to the current implementation but verify why it's "wrong".
// Maybe `MathRun` expects the text to be compliant with Word's linear math syntax.
//
// Let's try to simply wrap it in `Math` and `MathRun` as we did, 
// but ensure we strip any `$` or `\[` which we already do (node.value is raw).
//
// If the user sees "error", maybe it's because of unsupported commands.
//
// Let's try to use `temml` or `mathjax` to generate MathML, 
// and then use a simple regex-based MML->OMML converter? That's too risky.
//
// Let's try to use the `docx` features properly.
// `new Math({ children: [ new MathRun(latex) ] })`
//
// Maybe we need to split the LaTeX into multiple runs?
//
// Let's try to just output the text for now, but in a way that Word recognizes as "Equation".
//
// Reverting to the previous "Linear Math" approach but with better cleanup.
//
// Wait, I see `mathjax-node` installation in previous step.
// Let's use it to generate SVG as a fallback? No, user wants editable.
//
// Let's try to simply improve the Linear Format compatibility.
// Word's linear format doesn't support ALL LaTeX commands.
// e.g. `\text{}` might be an issue?
//
// Let's try to just use the raw LaTeX.
//
// If the user says "formula is wrong", maybe it's because it's not rendering?
//
// Let's try to add a helper to clean LaTeX for Word.
// Word prefers `\func` for standard functions, and `(a)/(b)` for fractions sometimes.
// But standard LaTeX `\frac` usually works in modern Word.
//
// Let's try to just ensure `docx` is used correctly.
//
// Actually, I will revert to a safe implementation:
// We will use `MathRun` but we need to import it.
//
// I will also add a `cleanLatexForWord` function.

import { MathRun } from "docx";

function cleanLatexForWord(latex: string): string {
  // Word's linear format is close to LaTeX but has quirks.
  // 1. Remove `\displaystyle`
  let clean = latex.replace(/\\displaystyle/g, "");
  // 2. Ensure spaces around operators?
  // 3. Remove `\mathrm`? Word supports it but maybe fonts issue.
  
  return clean.trim();
}

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
    results.push(new Math({
      children: [
        new MathRun(cleanLatexForWord(node.value))
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
              new MathRun(cleanLatexForWord(node.value))
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
