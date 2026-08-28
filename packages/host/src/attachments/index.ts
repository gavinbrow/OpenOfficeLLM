// Attachment handling: on-disk store, text extraction, and OCR fallback.
//
// The pane uploads files via multipart POST to /api/attachments and receives
// back refs it can include on a ChatRequest. The host resolves those refs to
// extracted text (folded into the system prompt) or image content blocks
// (inlined into the first user message for vision-capable models) at chat
// time. This barrel keeps the three modules' public surfaces in one place so
// server.ts and the prompt builder import from a single entry.

export { saveAttachment, getMeta, getBytes, deleteAttachment, cleanup } from './store.js'
export { extractText, AttachmentError } from './extract.js'
export type { ExtractResult } from './extract.js'
export { extractTextOcr } from './ocr.js'
export type { AttachmentMeta, AttachmentKind } from './store.js'