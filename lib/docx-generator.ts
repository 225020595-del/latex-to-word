import {
  Document,
  HeadingLevel,
  Math,
  MathFraction,
  MathIntegral,
  type MathComponent,
  MathRadical,
  MathRun,
  MathSubScript,
  MathSubSuperScript,
  MathSum,
  MathSuperScript,
  Packer,
  Paragraph,
  TextRun,
} from "docx";
import { XMLParser } from "fast-xml-parser";
import temml from "temml";
import { unified } from "unified";
import remarkParse from "remark-parse";
import remarkMath from "remark-math";

type MathmlNode = {
  name: string;
  children: MathmlNode[];
  text?: string;
};

const mathmlParser = new XMLParser({
  ignoreAttributes: true,
  preserveOrder: true,
  textNodeName: "#text",
});

function buildMathmlNodes(parsed: any): MathmlNode[] {
  if (!Array.isArray(parsed)) return [];
  const nodes: MathmlNode[] = [];

  for (const item of parsed) {
    if (!item || typeof item !== "object") continue;
    const keys = Object.keys(item);
    if (keys.length === 0) continue;
    const key = keys[0];
    const value = (item as any)[key];

    if (key === "#text") {
      const text = typeof value === "string" ? value : "";
      if (text.trim() !== "") nodes.push({ name: "#text", children: [], text });
      continue;
    }

    const children = Array.isArray(value) ? buildMathmlNodes(value) : [];
    const text = typeof value === "string" ? value : undefined;
    nodes.push({ name: key, children, text });
  }

  return nodes;
}

function getMathmlText(node: MathmlNode): string {
  if (node.name === "#text") return node.text ?? "";
  if (node.text) return node.text;
  return node.children.map(getMathmlText).join("");
}

function toComponents(node: MathmlNode): MathComponent[] {
  if (node.name === "math" || node.name === "mrow" || node.name === "mstyle") return seqToComponents(node.children);
  return nodeToComponents(node);
}

function seqToComponents(nodes: MathmlNode[]): MathComponent[] {
  const out: MathComponent[] = [];

  for (let i = 0; i < nodes.length; i++) {
    const node = nodes[i];

    if (node.name === "munderover" || node.name === "munder" || node.name === "mover") {
      const baseNode = node.children[0];
      const underNode = node.children[1];
      const overNode = node.children[2];
      const baseText = baseNode ? getMathmlText(baseNode).trim() : "";
      const subScript = underNode ? toComponents(underNode) : undefined;
      const superScript = overNode ? toComponents(overNode) : undefined;

      const next = i + 1 < nodes.length ? nodes[i + 1] : undefined;
      if (baseText.includes("∑") && next) {
        out.push(new MathSum({ children: toComponents(next), subScript, superScript }));
        i++;
        continue;
      }

      if (baseText.includes("∫") && next) {
        out.push(new MathIntegral({ children: toComponents(next), subScript, superScript }));
        i++;
        continue;
      }
    }

    out.push(...nodeToComponents(node));
  }

  return out;
}

function nodeToComponents(node: MathmlNode): MathComponent[] {
  if (node.name === "math" || node.name === "mrow" || node.name === "mstyle") return seqToComponents(node.children);

  if (node.name === "mi" || node.name === "mn" || node.name === "mo" || node.name === "mtext") {
    const text = getMathmlText(node).replace(/\s+/g, " ").trim();
    return text ? [new MathRun(text)] : [];
  }

  if (node.name === "mspace") {
    return [];
  }

  if (node.name === "mfrac") {
    const numerator = node.children[0] ? toComponents(node.children[0]) : [];
    const denominator = node.children[1] ? toComponents(node.children[1]) : [];
    return [new MathFraction({ numerator, denominator })];
  }

  if (node.name === "msqrt") {
    return [new MathRadical({ children: seqToComponents(node.children) })];
  }

  if (node.name === "mroot") {
    const base = node.children[0] ? toComponents(node.children[0]) : [];
    const degree = node.children[1] ? toComponents(node.children[1]) : [];
    return [new MathRadical({ children: base, degree })];
  }

  if (node.name === "msup") {
    const base = node.children[0] ? toComponents(node.children[0]) : [];
    const superScript = node.children[1] ? toComponents(node.children[1]) : [];
    return [new MathSuperScript({ children: base, superScript })];
  }

  if (node.name === "msub") {
    const base = node.children[0] ? toComponents(node.children[0]) : [];
    const subScript = node.children[1] ? toComponents(node.children[1]) : [];
    return [new MathSubScript({ children: base, subScript })];
  }

  if (node.name === "msubsup") {
    const base = node.children[0] ? toComponents(node.children[0]) : [];
    const subScript = node.children[1] ? toComponents(node.children[1]) : [];
    const superScript = node.children[2] ? toComponents(node.children[2]) : [];
    return [new MathSubSuperScript({ children: base, subScript, superScript })];
  }

  if (node.name === "mfenced") {
    return seqToComponents(node.children);
  }

  if (node.children.length > 0) {
    return seqToComponents(node.children);
  }

  const text = getMathmlText(node).replace(/\s+/g, " ").trim();
  return text ? [new MathRun(text)] : [];
}

function latexToMathComponents(latex: string, displayMode: boolean): MathComponent[] {
  const mathml = temml.renderToString(latex, { displayMode, xml: true });
  const parsed = mathmlParser.parse(mathml);
  const nodes = buildMathmlNodes(parsed);
  const mathNode = nodes.find((n) => n.name === "math");
  if (!mathNode) return [new MathRun(latex)];
  return seqToComponents(mathNode.children);
}

function processNode(node: any, parentStyle?: { bold?: boolean; italics?: boolean }): (TextRun | Math)[] {
  const results: (TextRun | Math)[] = [];
  
  if (node.type === "text") {
    results.push(new TextRun({ 
      text: node.value, 
      bold: parentStyle?.bold, 
      italics: parentStyle?.italics 
    }));
  } else if (node.type === "inlineMath") {
    results.push(new Math({ children: latexToMathComponents(node.value, false) }));
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
      children.push(
        new Paragraph({
          children: [new Math({ children: latexToMathComponents(node.value, true) })],
          alignment: "center",
          spacing: { before: 200, after: 200 },
        })
      );
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
