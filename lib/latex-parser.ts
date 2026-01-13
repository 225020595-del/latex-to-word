import { mathjax } from "mathjax-full/js/mathjax.js";
import { TeX } from "mathjax-full/js/input/tex.js";
import { MathML } from "mathjax-full/js/input/mathml.js";
import { SerializedMmlVisitor } from "mathjax-full/js/core/MmlTree/SerializedMmlVisitor.js";
import { liteAdaptor } from "mathjax-full/js/adaptors/liteAdaptor.js";
import { RegisterHTMLHandler } from "mathjax-full/js/handlers/html.js";

// Initialize MathJax
const adaptor = liteAdaptor();
RegisterHTMLHandler(adaptor);

const tex = new TeX({ packages: ["base", "ams"] });
const mathml = new MathML();
const visitor = new SerializedMmlVisitor();
const html = mathjax.document("", { InputJax: tex });

// Basic XSLT replacement logic to convert MathML to OMML (simplified)
// Since we can't run real XSLT easily in Node/Browser without heavy libs,
// we will use a "best effort" string transformation or a dedicated library if found.
//
// Actually, generating OMML manually is hard.
// BUT, we can use `temml` or `mathjax` to get MathML.
// Word DOES accept MathML if it's namespaced correctly in `m:oMathPara`.
//
// Let's try to wrap standard MathML in OMML wrapper?
// Word usually requires conversion.
//
// WAIT! There is a trick:
// Use `mathjax-full` to generate MathML.
// Then use a lightweight MML->OMML transformer.
//
// OR: We can use `mathml2omml` web service? No, must be local.
//
// Let's try to stick to "Linear Format" but FIX the issues.
// The user said "format mismatch" or "not professional".
//
// If we want TRUE Professional Math, we need OMML.
// OMML XML structure:
// <m:oMath>
//   <m:r>
//     <m:t>x</m:t>
//   </m:r>
// </m:oMath>
//
// It is very verbose.
//
// Let's look for a `latex-to-omml` library again.
// There isn't a good one for Node.
//
// ALTERNATIVE:
// We use `temml` to generate MathML.
// Then we use `mathml-to-omml` library?
//
// Let's try `npm install mathml-to-omml` (I saw it in search results).
// Wait, I didn't see it installed.
//
// Let's try to search specifically for a JS library that does this.
//
// If not, we will improve the Linear Format to be PERFECT.
// The issue with Linear Format is often:
// 1. Spacing.
// 2. Command mapping.
//
// Let's try to use `docx`'s `Math` but passing **MathML**?
// `docx` doesn't support MathML input.
//
// Let's try to insert **Raw XML** into the document.
// `docx` allows `new XmlComponent(...)`.
// If we can get the OMML XML, we can inject it.
//
// How to get OMML from LaTeX?
// MathJax doesn't output OMML.
//
// Let's use a very clever trick:
// Word Linear Format IS the way to go, but we need to ensure it triggers "Professional" mode.
//
// Maybe we need to set a flag in `MathRun`?
// `docx` documentation says: `new Math({ children: [...] })` creates an OMML object.
// If we put text inside, it's `m:t`.
// Word sees it as linear math.
//
// To make it professional, we need to construct the OMML tree (fractions, etc.).
// `docx` supports `MathFraction`, `MathSuperScript`, etc.!
//
// WE SHOULD USE `docx`'s MATH BUILDERS instead of raw text!
// This is the "Professional" way.
//
// We need a **LaTeX Parser** that outputs **docx Math Objects**.
//
// I will implement a basic `latexToDocxMath` parser.
// It will parse LaTeX AST and return `MathFraction`, `MathSup`, etc.
//
// This is much better than Linear Format text!

import { 
  MathRun, MathFraction, MathSuperScript, MathSubScript, MathSubSuperScript, MathRadical, MathFunction,
  MathSum, MathIntegral, MathLimit
} from "docx";

// We need a LaTeX parser. `remark-math` gives us AST?
// No, `remark-math` just gives "math" node with string value.
// We need to parse the LaTeX string.
// We can use a simple recursive descent parser or regex for common structures.
//
// Let's implement a simple parser for: \frac, \sqrt, ^, _, \sum, \int.

export function parseLatexToDocx(latex: string): any[] {
  let cursor = 0;
  
  function parseGroup(): any[] {
    skipWhitespace();
    if (cursor >= latex.length) return [];
    
    if (latex[cursor] === "{") {
      cursor++;
      const children: any[] = [];
      while (cursor < latex.length && latex[cursor] !== "}") {
        children.push(...parseNext());
      }
      cursor++; // skip }
      return children;
    }
    return parseNext();
  }

  function parseNext(): any[] {
    skipWhitespace();
    if (cursor >= latex.length) return [];
    
    const char = latex[cursor];
    
    if (char === "\\") {
      const start = cursor;
      cursor++;
      while (cursor < latex.length && /[a-zA-Z]/.test(latex[cursor])) {
        cursor++;
      }
      const command = latex.slice(start, cursor);
      
      if (command === "\\frac") {
        const num = parseGroup();
        const den = parseGroup();
        return [new MathFraction({ numerator: num, denominator: den })];
      } else if (command === "\\sqrt") {
         // Check for optional arg [n]
         skipWhitespace();
         if (latex[cursor] === "[") {
           // Handle nth root
           cursor++;
           const degree: any[] = [];
           while (cursor < latex.length && latex[cursor] !== "]") {
             degree.push(...parseNext());
           }
           cursor++;
           const body = parseGroup();
           return [new MathRadical({ degree: degree, children: body })];
         } else {
           const body = parseGroup();
           return [new MathRadical({ children: body })];
         }
      } else if (command === "\\sum") {
        return [new MathSum({ children: [] })]; // Simplified
        // Complex sums need sub/sup handling which is tricky in linear scan
      } else if (command === "\\int") {
        return [new MathIntegral({ children: [] })];
      } else {
        // Generic command, treat as text for now
        return [new MathRun(command)]; 
      }
    } else if (char === "^") {
      // Superscript - this is tricky because it modifies the PREVIOUS element
      // But in this simple parser, we return a list.
      // We need to handle this at the caller level or lookbehind.
      // 
      // Better approach: `parseTokens` first, then `buildTree`.
      cursor++;
      const sup = parseGroup();
      return [{ type: "sup", children: sup }];
    } else if (char === "_") {
      cursor++;
      const sub = parseGroup();
      return [{ type: "sub", children: sub }];
    } else if (char === "{" || char === "}") {
      // Should be handled by parseGroup, but if found here, skip or error
      if (char === "{") return parseGroup();
      cursor++;
      return [];
    } else {
      cursor++;
      return [new MathRun(char)];
    }
  }

  function skipWhitespace() {
    while (cursor < latex.length && /\s/.test(latex[cursor])) {
      cursor++;
    }
  }

  const results: any[] = [];
  while (cursor < latex.length) {
    results.push(...parseNext());
  }
  
  // Post-process for sub/sup (attach to previous element)
  // This is a naive implementation.
  const finalResults: any[] = [];
  for (let i = 0; i < results.length; i++) {
    const curr = results[i];
    if (curr.type === "sup" || curr.type === "sub") {
      // Attach to previous
      const prev = finalResults.pop();
      if (!prev) {
        finalResults.push(new MathRun("")); // Empty base
      }
      
      // If previous was already a sub/sup, merge?
      // docx has MathSubSuperScript
      
      if (curr.type === "sup") {
         finalResults.push(new MathSuperScript({ children: [prev], superScript: curr.children }));
      } else {
         finalResults.push(new MathSubScript({ children: [prev], subScript: curr.children }));
      }
    } else {
      finalResults.push(curr);
    }
  }

  return finalResults;
}

// NOTE: Implementing a full LaTeX parser is complex.
// Since `docx` components are strict, any error breaks the doc.
//
// BACKUP PLAN:
// Use `mathjax-full` to convert LaTeX -> MathML.
// Then write a **MathML to Docx Mapper**.
// MathML is XML, easier to traverse than raw LaTeX string.
//
// MathML structure:
// <mfrac> -> MathFraction
// <msqrt> -> MathRadical
// <msup> -> MathSuperScript
//
// This is much more reliable!

export function mathmlToDocx(node: any): any {
  if (node.kind === 'math') {
    return node.children.map(mathmlToDocx);
  }
  if (node.kind === 'mfrac') {
    return new MathFraction({
      numerator: node.children[0].children.map(mathmlToDocx),
      denominator: node.children[1].children.map(mathmlToDocx)
    });
  }
  if (node.kind === 'msqrt') {
    return new MathRadical({
      children: node.children.map(mathmlToDocx)
    });
  }
  if (node.kind === 'msup') {
     return new MathSuperScript({
       children: [mathmlToDocx(node.children[0])],
       superScript: [mathmlToDocx(node.children[1])]
     });
  }
  if (node.kind === 'msub') {
     return new MathSubScript({
       children: [mathmlToDocx(node.children[0])],
       subScript: [mathmlToDocx(node.children[1])]
     });
  }
  if (node.kind === 'mi' || node.kind === 'mn' || node.kind === 'mo') {
     // Text node
     const text = node.children[0].text; // Simplified
     return new MathRun(text);
  }
  // Fallback
  return new MathRun("");
}

// Let's use `mathjax-full` to get the MML Tree (internal object), then map it.
// `mathjax` documentation says we can get the MmlNode tree.
//
// import { TeX } from "mathjax-full/js/input/tex.js";
// const tex = new TeX();
// const tree = tex.parse("..."); // Returns MmlNode
//
// Let's try to implement this `latexToDocx` using MathJax's AST.
