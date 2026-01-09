export function normalizeLatex(text: string): string {
  // Replace block math \[ ... \] with $$ ... $$
  // We use a regex that matches \[ followed by anything until \]
  // s flag allows . to match newlines
  let normalized = text.replace(/\\\[([\s\S]*?)\\\]/g, (_, match) => {
    return `$$${match}$$`;
  });

  // Replace inline math \( ... \) with $ ... $
  // We need to be careful not to break escaped parenthesis if any (though unlikely in this context)
  normalized = normalized.replace(/\\\(([\s\S]*?)\\\)/g, (_, match) => {
    return `$${match}$`;
  });

  return normalized;
}
