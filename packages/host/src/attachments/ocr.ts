// OCR fallback for image attachments whose text matters more than their
// pixels — a scanned PDF rendered to PNG, a screenshot of a code editor, a
// photographed whiteboard. The primary image path forwards the bytes as a
// vision content block to a vision-capable model; OCR is the backstop for
// models that cannot see, and for the case where the user wants the text in
// the prompt rather than the image.
//
// tesseract.js ships a ~10MB worker and language data, downloaded on first
// use and cached under `cache/tesseract` so repeat runs are offline. The
// worker is created and terminated per call rather than pooled: OCR is rare
// and latency-sensitive only relative to a single attachment, and a pooled
// worker would hold the WASM module resident for the whole host lifetime.

import { TESSERACT_CACHE_DIR } from '../paths.js'

export async function extractTextOcr(buffer: Buffer, _mimeType: string): Promise<string> {
  // Lazy-load so the ~10MB worker only downloads when OCR is actually needed.
  const { createWorker } = await import('tesseract.js')
  const worker = await createWorker('eng', 1, {
    cacheMethod: 'write',
    cachePath: TESSERACT_CACHE_DIR,
    logger: () => {}, // silence progress spam
  })
  try {
    const { data } = await worker.recognize(buffer)
    return data.text ?? ''
  } finally {
    await worker.terminate()
  }
}