"use server";

import { createDocxFromMarkdown } from "@/lib/docx-generator";

export async function generateDocx(markdown: string): Promise<string> {
  try {
    // Now we pass the raw markdown instead of HTML
    return await createDocxFromMarkdown(markdown);
  } catch (error) {
    console.error("Error generating DOCX:", error);
    throw new Error("Failed to generate DOCX file");
  }
}
