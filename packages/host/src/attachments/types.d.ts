// Ambient module declarations for attachment extractor dependencies that
// ship no TypeScript types at the subpath the extractor imports.
//
// pdfjs-dist's main package types cover the `pdfjs-dist` entry, but the
// `legacy/build/pdf.worker.mjs` subpath — the Node-friendly worker module we
// inject as `globalThis.pdfjsWorker` so the fake-worker loader does not try
// to `import("./pdf.worker.mjs")` a path the CJS bundle does not have — has
// no published type declaration. Declaring only the one export we read keeps
// the surface narrow; a future real `@types/pdfjs-dist` subpath only has to
// match `WorkerMessageHandler` to take over.
declare module 'pdfjs-dist/legacy/build/pdf.worker.mjs' {
  /** The worker-side message handler pdf.js's fake-worker loader reads off
   *  `globalThis.pdfjsWorker`. We only ever assign the module namespace to
   *  that global; the shape below is the one field the loader touches. */
  export const WorkerMessageHandler: unknown
}