// pdfjs-dist ships without clean ESM type resolution for this Node subpath,
// so we declare just the bits this project uses.
declare module "pdfjs-dist/legacy/build/pdf.mjs" {
  export function getDocument(params: { data: Uint8Array }): {
    promise: Promise<{
      numPages: number;
      getPage: (pageNumber: number) => Promise<{
        getTextContent: () => Promise<{ items: unknown[] }>;
        getViewport: (params: { scale: number }) => { width: number; height: number };
      }>;
    }>;
  };
}
