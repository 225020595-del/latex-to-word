// Simple parser to convert LaTeX to Word's Linear Math format
// Word Linear Format Reference: http://www.unicode.org/notes/tn28/UTN28-PlainTextMath-v3.pdf

export function latexToWordLinear(latex: string): string {
  let cursor = 0;
  
  // Helper to consume the next group {...} or single token
  function consumeGroup(): string {
    skipWhitespace();
    if (cursor >= latex.length) return "";
    
    if (latex[cursor] === "{") {
      cursor++; // skip {
      let depth = 1;
      let start = cursor;
      while (cursor < latex.length && depth > 0) {
        if (latex[cursor] === "{") depth++;
        else if (latex[cursor] === "}") depth--;
        cursor++;
      }
      // Return content without braces
      return latexToWordLinear(latex.slice(start, cursor - 1));
    } else {
      // Consume single char or command
      if (latex[cursor] === "\\") {
        let start = cursor;
        cursor++;
        while (cursor < latex.length && /[a-zA-Z]/.test(latex[cursor])) {
          cursor++;
        }
        return latex.slice(start, cursor);
      } else {
        return latex[cursor++];
      }
    }
  }

  function skipWhitespace() {
    while (cursor < latex.length && /\s/.test(latex[cursor])) {
      cursor++;
    }
  }

  let result = "";
  
  while (cursor < latex.length) {
    const char = latex[cursor];

    if (char === "\\") {
      // Command
      const start = cursor;
      cursor++;
      while (cursor < latex.length && /[a-zA-Z]/.test(latex[cursor])) {
        cursor++;
      }
      const command = latex.slice(start, cursor);

      if (command === "\\frac") {
        const numerator = consumeGroup();
        const denominator = consumeGroup();
        result += `(${numerator})/(${denominator})`;
      } else if (command === "\\sqrt") {
        // Check for optional argument [n]
        skipWhitespace();
        if (latex[cursor] === "[") {
          cursor++; // skip [
          let depth = 1;
          let argStart = cursor;
          while (cursor < latex.length && depth > 0) {
             if (latex[cursor] === "[") depth++;
             else if (latex[cursor] === "]") depth--;
             cursor++;
          }
          const rootDegree = latex.slice(argStart, cursor - 1);
          const body = consumeGroup();
          result += `\\sqrt(${rootDegree}&${body})`;
        } else {
          const body = consumeGroup();
          result += `\\sqrt(${body})`;
        }
      } else if (command === "\\left") {
         // Skip \left, just take the next char
         skipWhitespace();
         result += latex[cursor++];
      } else if (command === "\\right") {
         skipWhitespace();
         result += latex[cursor++];
      } else if (command === "\\text") {
         const text = consumeGroup();
         result += `"${text}"`;
      } else if (command === "\\displaystyle" || command === "\\limits" || command === "\\nolimits") {
         // Ignore these layout commands
      } else {
         // Keep other commands as is (e.g. \alpha, \int, \sum)
         result += command + " "; // Add space to ensure separation
      }
    } else if (char === "{" || char === "}") {
      // Skip outer braces that are not part of a command (grouping)
      // But we need to be careful. In Word, () is grouping.
      // If we see raw {}, it might be for grouping.
      // Let's assume consumeGroup handles command arguments, so raw {} are grouping.
      // Convert { to ( and } to )?
      // No, strictly, {} in LaTeX is grouping, in Word it's hidden grouping.
      // Word uses () for visible grouping or logical grouping.
      // Let's ignore them for now? Or map to ()?
      // If we map to (), it might show ().
      // Let's skip them, relying on the fact that we handled arguments already.
      cursor++;
    } else if (char === "^" || char === "_") {
      cursor++;
      const arg = consumeGroup();
      // Word needs () around complex arguments for sub/sup
      // If arg is simple (1 char), no need, but () is safer.
      result += `${char}(${arg})`;
    } else {
      result += char;
      cursor++;
    }
  }

  // Final cleanup
  // Fix double spaces
  result = result.replace(/\s+/g, " ");
  // Fix empty parenthesis () -> empty?
  
  return result.trim();
}
