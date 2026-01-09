declare module 'html-to-docx' {
  const htmlToDocx: (
    html: string,
    headerHTML?: string | null,
    documentOptions?: any,
    footerHTML?: string | null
  ) => Promise<Blob>;
  export default htmlToDocx;
}
