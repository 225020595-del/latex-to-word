import { createDocxFromMarkdown } from "./lib/docx-generator";
import * as fs from "fs";

async function test() {
  const markdown = `
# Test Document

Here is an inline equation: $E=mc^2$.

Block equation:
$$
\\int_0^\\infty x^2 dx
$$
`;

  try {
    console.log("Generating DOCX...");
    const base64 = await createDocxFromMarkdown(markdown);
    const buffer = Buffer.from(base64, "base64");
    fs.writeFileSync("test-output.docx", buffer);
    console.log("Done! Saved to test-output.docx");
  } catch (e) {
    console.error("Error:", e);
  }
}

test();
