import { MmlVisitor } from "mathjax-full/js/core/MmlTree/MmlVisitor.js";
import { MmlNode, TextNode, XMLNode } from "mathjax-full/js/core/MmlTree/MmlNode.js";

export class OmmlVisitor extends MmlVisitor {
  // OMML Namespace
  // We don't need to add xmlns to every element if we wrap the root properly, 
  // but docx xml component might need it. 
  // Usually <m:oMath> is enough if namespaces are defined in document.
  // But for safety, we just generate the tags with m: prefix.

  constructor() {
    super();
  }

  public visitTree(node: MmlNode): string {
    return this.visitNode(node, "");
  }

  public visitNode(node: MmlNode, ...args: any[]): string {
    const kind = node.kind;
    const handler = (this as any)["visit" + kind.charAt(0).toUpperCase() + kind.slice(1)];
    if (handler) {
      return handler.call(this, node, ...args);
    }
    // Default fallback: visit children
    return this.visitChildren(node, ...args);
  }

  public visitChildren(node: MmlNode, ...args: any[]): string {
    let result = "";
    for (const child of node.childNodes) {
      result += this.visitNode(child as MmlNode, ...args);
    }
    return result;
  }

  // --- Math Nodes Handlers ---

  public visitMath(node: MmlNode, ...args: any[]): string {
    // Root node
    return `<m:oMath xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math">${this.visitChildren(node, ...args)}</m:oMath>`;
  }

  public visitMfrac(node: MmlNode, ...args: any[]): string {
    const num = this.visitNode(node.childNodes[0] as MmlNode, ...args);
    const den = this.visitNode(node.childNodes[1] as MmlNode, ...args);
    return `<m:f><m:num>${num}</m:num><m:den>${den}</m:den></m:f>`;
  }

  public visitMsqrt(node: MmlNode, ...args: any[]): string {
    const base = this.visitChildren(node, ...args);
    return `<m:rad><m:radPr><m:degHide m:val="on"/></m:radPr><m:e>${base}</m:e></m:rad>`;
  }

  public visitMroot(node: MmlNode, ...args: any[]): string {
    const base = this.visitNode(node.childNodes[0] as MmlNode, ...args);
    const degree = this.visitNode(node.childNodes[1] as MmlNode, ...args);
    return `<m:rad><m:deg>${degree}</m:deg><m:e>${base}</m:e></m:rad>`;
  }

  public visitMsup(node: MmlNode, ...args: any[]): string {
    const base = this.visitNode(node.childNodes[0] as MmlNode, ...args);
    const sup = this.visitNode(node.childNodes[1] as MmlNode, ...args);
    return `<m:sSup><m:e>${base}</m:e><m:sup>${sup}</m:sup></m:sSup>`;
  }

  public visitMsub(node: MmlNode, ...args: any[]): string {
    const base = this.visitNode(node.childNodes[0] as MmlNode, ...args);
    const sub = this.visitNode(node.childNodes[1] as MmlNode, ...args);
    return `<m:sSub><m:e>${base}</m:e><m:sub>${sub}</m:sub></m:sSub>`;
  }

  public visitMsubsup(node: MmlNode, ...args: any[]): string {
    const base = this.visitNode(node.childNodes[0] as MmlNode, ...args);
    const sub = this.visitNode(node.childNodes[1] as MmlNode, ...args);
    const sup = this.visitNode(node.childNodes[2] as MmlNode, ...args);
    return `<m:sSubSup><m:e>${base}</m:e><m:sub>${sub}</m:sub><m:sup>${sup}</m:sup></m:sSubSup>`;
  }

  // Token Elements
  public visitMi(node: MmlNode, ...args: any[]): string {
    return this.createRun((node as any).texClass === undefined ? "norm" : "math", this.getText(node));
  }

  public visitMn(node: MmlNode, ...args: any[]): string {
    return this.createRun("norm", this.getText(node));
  }

  public visitMo(node: MmlNode, ...args: any[]): string {
    const text = this.getText(node);
    // Special handling for operators could be added here
    return this.createRun("math", text);
  }

  public visitMtext(node: MmlNode, ...args: any[]): string {
    return this.createRun("norm", this.getText(node));
  }
  
  public visitMspace(node: MmlNode, ...args: any[]): string {
      return ""; // Ignore spaces for now or convert to m:t space?
  }

  // Layouts
  public visitMrow(node: MmlNode, ...args: any[]): string {
    return this.visitChildren(node, ...args);
  }
  
  public visitMstyle(node: MmlNode, ...args: any[]): string {
      return this.visitChildren(node, ...args);
  }

  // Advanced: Sums and Integrals (munderover, munder, mover)
  // MathJax often parses \sum as munderover or msubsup depending on displaystyle
  public visitMunderover(node: MmlNode, ...args: any[]): string {
    const base = this.visitNode(node.childNodes[0] as MmlNode, ...args);
    const under = this.visitNode(node.childNodes[1] as MmlNode, ...args);
    const over = this.visitNode(node.childNodes[2] as MmlNode, ...args);
    
    // Check if it's a large operator (sum/int)
    // We assume it is for now, or check base content
    return `<m:nary><m:naryPr><m:limLoc m:val="undOvr"/></m:naryPr><m:sub>${under}</m:sub><m:sup>${over}</m:sup><m:e>${base}</m:e></m:nary>`;
  }
  
  public visitMunder(node: MmlNode, ...args: any[]): string {
      const base = this.visitNode(node.childNodes[0] as MmlNode, ...args);
      const under = this.visitNode(node.childNodes[1] as MmlNode, ...args);
      // Assuming limits behavior
      return `<m:nary><m:naryPr><m:limLoc m:val="undOvr"/><m:supHide m:val="on"/></m:naryPr><m:sub>${under}</m:sub><m:sup></m:sup><m:e>${base}</m:e></m:nary>`;
  }
  
  public visitMover(node: MmlNode, ...args: any[]): string {
      const base = this.visitNode(node.childNodes[0] as MmlNode, ...args);
      const over = this.visitNode(node.childNodes[1] as MmlNode, ...args);
      return `<m:nary><m:naryPr><m:limLoc m:val="undOvr"/><m:subHide m:val="on"/></m:naryPr><m:sub></m:sub><m:sup>${over}</m:sup><m:e>${base}</m:e></m:nary>`;
  }

  // Helpers
  private getText(node: MmlNode): string {
    if (node.childNodes.length > 0 && node.childNodes[0].kind === "text") {
      return (node.childNodes[0] as TextNode).getText();
    }
    return "";
  }

  private createRun(style: string, text: string): string {
    // Style can be used to set m:scr etc.
    // For simplicity, we just return m:r > m:t
    return `<m:r><m:t>${text}</m:t></m:r>`;
  }
}
