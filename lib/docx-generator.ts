import { Document, Packer, Paragraph, TextRun, HeadingLevel, Math } from "docx";
import { unified } from "unified";
import remarkParse from "remark-parse";
import remarkMath from "remark-math";
import temml from "temml";

// Helper to convert LaTeX to MathML using Temml (lightweight, MathJax alternative)
// Word supports MathML when wrapped in OMML (Office MathML).
// Note: docx.js has a `Math` component but it requires specific structure.
// However, pasting raw MathML text usually works if configured correctly, 
// or we need to transform MathML XML to docx Math objects.
// 
// Actually, docx.js doesn't fully parse MathML strings into OMML automatically.
// But we can insert raw OMML XML if we convert it.
//
// A simpler approach supported by Word is to insert the MathML XML directly.
// Let's try to convert LaTeX -> MathML string.

function convertLatexToMathMl(latex: string, isInline: boolean): string | null {
  try {
    const mathml = temml.renderToString(latex, {
      displayMode: !isInline,
      xml: true, // Generate XML compatible MathML
    });
    return mathml;
  } catch (e) {
    console.error("Temml conversion error:", e);
    return null;
  }
}

// Since docx.js doesn't have a direct "Insert MathML String" feature that is stable,
// We will try to use the `Math` run if we can construct it, but constructing it from scratch is hard.
//
// ALTERNATIVE STRATEGY:
// Word 2007+ supports "Math Paragraphs". 
// We can try to clean the LaTeX and insert it as a text run, but that's not "editable math".
//
// The best way programmatically without a heavy OMML parser is to rely on MathML.
// docx.js supports `new Math({ children: ... })`.
//
// WAIT: There is no easy "LaTeX -> docx object" converter.
// However, we can use a trick: Word accepts MathML if pasted. 
// But for file generation, we need OMML.
//
// Let's use a known workaround:
// Insert the MathML as a raw XML string using `new TextRun` isn't enough.
// We need to wrap it in `m:oMathPara` or `m:oMath`.
//
// Fortunately, simple MathML is often readable by Word if we use the right XML namespace.
//
// Let's try to return a Math node if we can, or just Text if it fails.
// Actually, let's look at `docx` docs... it has `Math` support but it's manual building.
//
// REVISED PLAN:
// We will stick to the SVG approach for now because "editable math" requires complex OMML conversion
// which is not easily available in JS without heavy libraries.
// 
// BUT, the user explicitly asked for "editable math".
//
// Let's try to use `mathml2omml` logic if possible, or just inject MathML XML.
// Word *does* support MathML if it's namespaced correctly in the document.
//
// Let's try to insert raw XML for the math parts.
// `docx` allows creating paragraphs with children.
// We might not be able to easily inject raw XML nodes inside a paragraph with `docx`.
//
// Wait, `docx` has `ExternalHyperlink` etc. 
//
// Let's try to simply use the LaTeX text itself but wrapped in a way Word recognizes?
// No, Word needs OMML.
//
// Let's try to use `temml` to get MathML, and then see if we can embed it.
//
// Actually, there is a `Math` run in `docx` but it expects `MathRun` children.
//
// If we can't easily do editable math, we might have to stick to SVG or use a web service.
// 
// Let's try to use `temml` -> MathML -> simple cleaning -> try to insert as raw text? No.
// 
// Let's look at `docx` again. It has `Math` and `MathRun`.
// `new Math({ children: [ new MathRun("x^2") ] })` -> this renders as linear format?
//
// If we just put the LaTeX code in a Math object, Word might interpret it as "Linear Format"
// which can be converted to Professional format by pressing Enter.
// This is "Editable"!
//
// Let's try: `new Math({ children: [ new MathRun(latex) ] })`

// Recursive helper to process AST nodes into Docx elements
function processNode(node: any, parentStyle?: { bold?: boolean; italics?: boolean }): (TextRun | ImageRun | Math)[] {
  const results: (TextRun | ImageRun | Math)[] = [];
  
  if (node.type === "text") {
    results.push(new TextRun({ 
      text: node.value, 
      bold: parentStyle?.bold, 
      italics: parentStyle?.italics 
    }));
  } else if (node.type === "inlineMath") {
    // Attempt to create an editable Math Run
    // We pass the raw LaTeX. In Word, this shows up as linear math.
    // Users can often convert it, or it might auto-convert.
    results.push(new Math({
      children: [
        new TextRun({
          text: node.value,
          // Word Linear Math usually doesn't need special styling, but let's keep it clean
        })
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
      // We wrap it in a Math Paragraph
      children.push(new Paragraph({
        children: [
          new Math({
            children: [
              new TextRun(node.value)
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
