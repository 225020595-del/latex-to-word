"use client";

import React, { useState } from "react";
import ReactMarkdown from "react-markdown";
import remarkMath from "remark-math";
import rehypeKatex from "rehype-katex";
import { Download, FileText, RefreshCw } from "lucide-react";
import { saveAs } from "file-saver";
import "katex/dist/katex.min.css";
import { generateDocx } from "@/app/actions";
import { normalizeLatex } from "@/lib/utils";

export default function Converter() {
  const [input, setInput] = useState<string>(
    "# Math Export Demo\n\nHere is an inline equation: $E=mc^2$.\n\nAnd here is a block equation:\n\n$$\n\\int_{-\\infty}^{\\infty} e^{-x^2} dx = \\sqrt{\\pi}\n$$\n\nTry copying this into Word!"
  );
  const [isExporting, setIsExporting] = useState(false);

  // Normalized input for rendering and exporting
  // We process the input on the fly to support \( ... \) and \[ ... \]
  const normalizedInput = normalizeLatex(input);

  // Function to handle Docx Export
  const handleExport = async () => {
    setIsExporting(true);
    try {
      // Use the normalized input for export as well
      const base64Data = await generateDocx(normalizedInput);

      // Convert Base64 to Blob
      const response = await fetch(
        `data:application/vnd.openxmlformats-officedocument.wordprocessingml.document;base64,${base64Data}`
      );
      const blob = await response.blob();

      // Save the file
      saveAs(blob, "converted-math.docx");
    } catch (error) {
      console.error("Export failed:", error);
      alert("Export failed. See console for details.");
    } finally {
      setIsExporting(false);
    }
  };

  return (
    <div className="min-h-screen bg-gray-50 flex flex-col font-sans text-gray-900">
      {/* Header */}
      <header className="bg-white border-b border-gray-200 px-6 py-4 flex items-center justify-between sticky top-0 z-10 shadow-sm">
        <div className="flex items-center gap-2">
          <div className="bg-indigo-600 p-2 rounded-lg text-white">
            <FileText size={20} />
          </div>
          <h1 className="text-xl font-bold text-gray-800">
            AI Latex to Word Converter
          </h1>
        </div>
        <button
          onClick={handleExport}
          disabled={isExporting}
          className="flex items-center gap-2 bg-indigo-600 hover:bg-indigo-700 text-white px-4 py-2 rounded-md font-medium transition-colors disabled:opacity-50 disabled:cursor-not-allowed shadow-sm"
        >
          {isExporting ? (
            <RefreshCw className="animate-spin" size={18} />
          ) : (
            <Download size={18} />
          )}
          {isExporting ? "Generating..." : "Export to .docx"}
        </button>
      </header>

      {/* Main Content */}
      <main className="flex-1 flex flex-col md:flex-row h-[calc(100vh-64px)] overflow-hidden">
        {/* Left Column: Input */}
        <div className="w-full md:w-1/2 flex flex-col border-r border-gray-200 bg-white">
          <div className="bg-gray-100 px-4 py-2 border-b border-gray-200 text-sm font-medium text-gray-500 uppercase tracking-wide">
            Markdown Input (LaTeX supported)
          </div>
          <textarea
            value={input}
            onChange={(e) => setInput(e.target.value)}
            className="flex-1 p-6 resize-none focus:outline-none font-mono text-sm leading-relaxed"
            placeholder="Type your markdown here..."
            spellCheck={false}
          />
        </div>

        {/* Right Column: Preview */}
        <div className="w-full md:w-1/2 flex flex-col bg-gray-50">
          <div className="bg-gray-100 px-4 py-2 border-b border-gray-200 text-sm font-medium text-gray-500 uppercase tracking-wide flex justify-between items-center">
            <span>Live Preview</span>
            <span className="text-xs text-gray-400 normal-case flex items-center gap-1">
              Powered by <span className="font-semibold text-gray-500">KaTeX</span>
            </span>
          </div>
          <div className="flex-1 p-8 overflow-auto prose prose-indigo max-w-none prose-img:rounded-lg prose-headings:font-bold prose-p:leading-relaxed prose-pre:bg-gray-800 prose-pre:text-white">
            <ReactMarkdown
              remarkPlugins={[remarkMath]}
              rehypePlugins={[rehypeKatex]}
            >
              {normalizedInput}
            </ReactMarkdown>
          </div>
        </div>
      </main>
    </div>
  );
}
