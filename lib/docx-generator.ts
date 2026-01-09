import { Document, Packer, Paragraph, TextRun, ImageRun } from "docx";
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
    
    // adaptor.innerHTML returns the container HTML, e.g., <mjx-container ...><svg ...>...</svg></mjx-container>
    const fullHtml = adaptor.innerHTML(node);
    
    // We need to extract the raw SVG string
    // Simple regex to find <svg...>...</svg>
    const svgMatch = fullHtml.match(/<svg[\s\S]*?<\/svg>/);
    
    if (!svgMatch) {
      console.warn("No SVG found in MathJax output:", fullHtml);
      return null;
    }
    
    const svgString = svgMatch[0];
    
    // Extract dimensions from SVG string
    // MathJax output format: width="X.Yex" height="A.Bex"
    // We assume 1ex ≈ 8px for calculation (adjust as needed for Word)
    const widthMatch = svgString.match(/width="([\d.]+)ex"/);
    const heightMatch = svgString.match(/height="([\d.]+)ex"/);
    
    let width = 100;
    let height = 30;
    
    // Convert ex to px (approximate)
    // In Word, we might want to scale this up slightly for better visibility
    const EX_TO_PX = 10; // Slightly increased scaling for better visibility
    
    if (widthMatch && widthMatch[1]) {
      width = parseFloat(widthMatch[1]) * EX_TO_PX;
    }
    if (heightMatch && heightMatch[1]) {
      height = parseFloat(heightMatch[1]) * EX_TO_PX;
    }
    
    // Ensure minimum dimensions to avoid invisible 0-size images
    width = Math.max(width, 10);
    height = Math.max(height, 10);

    return { svg: svgString, width, height };
  } catch (e) {
    console.error("MathJax conversion error:", e);
    return null;
  }
}

export async function createDocxFromMarkdown(markdown: string): Promise<string> {
  // Parse Markdown
  const processor = unified().use(remarkParse).use(remarkMath);
  const ast = processor.parse(markdown);

  const children: (Paragraph)[] = [];

  // Simplified traversal
  for (const node of (ast as any).children) {
    if (node.type === "paragraph") {
      const runs: (TextRun | ImageRun)[] = [];
      
      for (const child of node.children) {
        if (child.type === "text") {
          runs.push(new TextRun(child.value));
        } else if (child.type === "inlineMath") {
           const result = convertLatexToSvg(child.value, true);
           if (result) {
             runs.push(
               new ImageRun({
                 data: Buffer.from(result.svg),
                 transformation: {
                   width: result.width,
                   height: result.height,
                 },
                 type: "svg",
                 fallback: {
                    data: FALLBACK_BUFFER,
                    type: "png",
                 }
               })
             );
           } else {
             runs.push(new TextRun({ text: `$${child.value}$`, color: "red" }));
           }
        } else if (child.type === "emphasis") {
           if (child.children?.[0]?.type === "text") {
             runs.push(new TextRun({ text: child.children[0].value, italics: true }));
           }
        } else if (child.type === "strong") {
           if (child.children?.[0]?.type === "text") {
             runs.push(new TextRun({ text: child.children[0].value, bold: true }));
           }
        }
      }
      children.push(new Paragraph({ children: runs }));

    } else if (node.type === "math") {
      // Block Math
      const result = convertLatexToSvg(node.value, false);
      if (result) {
        children.push(
          new Paragraph({
            children: [
              new ImageRun({
                data: Buffer.from(result.svg),
                transformation: {
                   width: result.width,
                   height: result.height,
                },
                type: "svg",
                fallback: {
                    data: FALLBACK_BUFFER,
                    type: "png",
                 }
              }),
            ],
            alignment: "center",
            spacing: { before: 200, after: 200 }, // Add some space around block math
          })
        );
      } else {
         children.push(new Paragraph({ children: [new TextRun({ text: `$$${node.value}$$`, color: "red" })] }));
      }
    } else if (node.type === "heading") {
      children.push(
        new Paragraph({
          text: node.children?.[0]?.value || "",
          heading: `Heading${node.depth}` as any,
        })
      );
    }
  }

  const doc = new Document({
    sections: [
      {
        properties: {},
        children: children,
      },
    ],
  });

  const buffer = await Packer.toBuffer(doc);
  return buffer.toString("base64");
}
