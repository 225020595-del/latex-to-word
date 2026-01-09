"use client";

import React, { useState } from "react";
import ReactMarkdown from "react-markdown";
import remarkMath from "remark-math";
import rehypeKatex from "rehype-katex";
import { Download, FileText, RefreshCw, Sparkles, Code2, Eye } from "lucide-react";
import { saveAs } from "file-saver";
import "katex/dist/katex.min.css";
import { generateDocx } from "@/app/actions";
import { normalizeLatex } from "@/lib/utils";

export default function Converter() {
  const [input, setInput] = useState<string>(
    "# Math Export Demo\n\nHere is an inline equation: $E=mc^2$.\n\nAnd here is a block equation:\n\n$$\n\\int_{-\\infty}^{\\infty} e^{-x^2} dx = \\sqrt{\\pi}\n$$\n\nTry copying this into Word!"
  );
  const [isExporting, setIsExporting] = useState(false);

  const normalizedInput = normalizeLatex(input);

  const handleExport = async () => {
    setIsExporting(true);
    try {
      const base64Data = await generateDocx(normalizedInput);
      const response = await fetch(
        `data:application/vnd.openxmlformats-officedocument.wordprocessingml.document;base64,${base64Data}`
      );
      const blob = await response.blob();
      saveAs(blob, "converted-math.docx");
    } catch (error) {
      console.error("Export failed:", error);
      alert("Export failed. See console for details.");
    } finally {
      setIsExporting(false);
    }
  };

  return (
    <div className="min-h-screen bg-slate-900 text-slate-100 font-sans flex flex-col overflow-hidden">
      {/* Decorative Background Elements */}
      <div className="fixed top-0 left-0 w-full h-full overflow-hidden -z-10 pointer-events-none">
        <div className="absolute top-[-20%] left-[-10%] w-[50%] h-[50%] rounded-full bg-indigo-600/20 blur-[120px]" />
        <div className="absolute bottom-[-20%] right-[-10%] w-[50%] h-[50%] rounded-full bg-violet-600/20 blur-[120px]" />
      </div>

      {/* Header */}
      <header className="px-6 py-4 flex items-center justify-between border-b border-white/10 backdrop-blur-md bg-slate-900/50 sticky top-0 z-20">
        <div className="flex items-center gap-3">
          <div className="bg-gradient-to-br from-indigo-500 to-violet-600 p-2.5 rounded-xl shadow-lg shadow-indigo-500/20">
            <Sparkles size={20} className="text-white" />
          </div>
          <div>
            <h1 className="text-xl font-bold bg-clip-text text-transparent bg-gradient-to-r from-white to-slate-400">
              AI Latex to Word
            </h1>
            <p className="text-xs text-slate-400 font-medium tracking-wide">
              SMART CONVERTER
            </p>
          </div>
        </div>
        <button
          onClick={handleExport}
          disabled={isExporting}
          className="group flex items-center gap-2 bg-white text-slate-900 px-5 py-2.5 rounded-xl font-semibold transition-all hover:bg-indigo-50 hover:shadow-lg hover:shadow-indigo-500/20 hover:scale-[1.02] active:scale-[0.98] disabled:opacity-50 disabled:cursor-not-allowed disabled:hover:scale-100"
        >
          {isExporting ? (
            <RefreshCw className="animate-spin text-indigo-600" size={18} />
          ) : (
            <Download className="text-indigo-600 group-hover:-translate-y-0.5 transition-transform" size={18} />
          )}
          {isExporting ? "Generating..." : "Export to .docx"}
        </button>
      </header>

      {/* Main Content */}
      <main className="flex-1 p-6 h-[calc(100vh-80px)]">
        <div className="flex flex-col md:flex-row gap-6 h-full max-w-[1920px] mx-auto">
          
          {/* Left Column: Input */}
          <div className="w-full md:w-1/2 flex flex-col bg-white rounded-2xl shadow-2xl overflow-hidden ring-1 ring-white/10 transition-all duration-300 hover:shadow-indigo-500/10">
            <div className="bg-slate-50/80 backdrop-blur-sm px-5 py-3 border-b border-slate-200 flex items-center gap-2">
              <Code2 size={16} className="text-indigo-500" />
              <span className="text-sm font-semibold text-slate-600 uppercase tracking-wider">
                Markdown / LaTeX Input
              </span>
            </div>
            <textarea
              value={input}
              onChange={(e) => setInput(e.target.value)}
              className="flex-1 p-6 resize-none focus:outline-none font-mono text-sm leading-relaxed text-slate-800 bg-white selection:bg-indigo-100"
              placeholder="Paste your AI-generated math content here..."
              spellCheck={false}
            />
          </div>

          {/* Right Column: Preview */}
          <div className="w-full md:w-1/2 flex flex-col bg-white rounded-2xl shadow-2xl overflow-hidden ring-1 ring-white/10 transition-all duration-300 hover:shadow-violet-500/10">
            <div className="bg-slate-50/80 backdrop-blur-sm px-5 py-3 border-b border-slate-200 flex items-center justify-between">
              <div className="flex items-center gap-2">
                <Eye size={16} className="text-violet-500" />
                <span className="text-sm font-semibold text-slate-600 uppercase tracking-wider">
                  Live Preview
                </span>
              </div>
              <span className="text-[10px] font-bold px-2 py-0.5 rounded-full bg-violet-100 text-violet-600">
                KaTeX Powered
              </span>
            </div>
            <div className="flex-1 p-8 overflow-auto bg-white">
              <div className="prose prose-slate max-w-none text-slate-800 prose-headings:text-slate-900 prose-headings:font-bold prose-p:text-slate-800 prose-p:leading-relaxed prose-pre:bg-slate-800 prose-pre:text-slate-50 prose-pre:rounded-xl prose-strong:text-slate-900 prose-li:text-slate-800">
                <ReactMarkdown
                  remarkPlugins={[remarkMath]}
                  rehypePlugins={[rehypeKatex]}
                >
                  {normalizedInput}
                </ReactMarkdown>
              </div>
            </div>
          </div>

        </div>
      </main>
    </div>
  );
}
