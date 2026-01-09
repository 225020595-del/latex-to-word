import { unified } from "unified";
import remarkParse from "remark-parse";
import remarkMath from "remark-math";
import remarkRehype from "remark-rehype";
import rehypeKatex from "rehype-katex";
import rehypeStringify from "rehype-stringify";

async function testRender() {
  const input = `
  公式： \\( MSE = \\frac{1}{n} \\)
  
  块级：
  \\[ E = mc^2 \\]
  `;

  const file = await unified()
    .use(remarkParse)
    .use(remarkMath)
    .use(remarkRehype)
    .use(rehypeKatex)
    .use(rehypeStringify)
    .process(input);

  console.log(String(file));
}

testRender();
