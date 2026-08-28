// On-disk attachment store keyed by host-assigned id.
//
// Each attachment lives at `ATTACHMENTS_DIR/<id>/<original-name>`. The id is a
// fresh UUID per upload, so two files with the same original name never
// collide and an attacker cannot guess another attachment's path. Only the
// original filename is preserved in the path — it is what gets sent back as
// `fileName` in the ref and shown in the UI, and it is what the prompt builder
// uses for the "the user attached a file named X" header.
//
// The bytes are kept on disk rather than in memory: a 50 MB PDF would balloon
// the host's resident set if held in a Map, and attachments are read back
// twice at most (once to extract text on upload, once when the pane requests
// the bytes for inline preview), so there is nothing to gain from caching
// them in RAM.
//
// The in-memory map holds metadata only: the on-disk path, the MIME type,
// the kind (text vs image), the original filename, a rough token estimate for
// the ContextChips budget, and — for text attachments — the already-extracted
// text so prompt assembly does not have to re-extract.

import fs from 'node:fs'
import path from 'node:path'
import crypto from 'node:crypto'
import { ATTACHMENTS_DIR } from '../paths.js'

export type AttachmentKind = 'text' | 'image'

export interface AttachmentMeta {
  id: string
  fileName: string
  mimeType: string
  kind: AttachmentKind
  tokenEstimate: number
  /** Extracted text, present only for text-kind attachments. Image
   *  attachments forward their bytes as vision content blocks instead. */
  text?: string
  /** Absolute path to the stored file. */
  path: string
}

interface SaveInput {
  name: string
  type: string
  buffer: Buffer
}

interface SaveResult {
  id: string
  fileName: string
  mimeType: string
  kind: AttachmentKind
}

const store = new Map<string, AttachmentMeta>()

/** Classify an upload as a text or image attachment from its MIME type. Anything
 *  starting with `image/` is forwarded as a vision content block; everything
 *  else is treated as a candidate for text extraction. */
function classifyKind(mimeType: string): AttachmentKind {
  return mimeType.toLowerCase().startsWith('image/') ? 'image' : 'text'
}

export function saveAttachment(file: SaveInput): SaveResult {
  const id = crypto.randomUUID()
  const kind = classifyKind(file.type)
  // `file.type` can be empty when the browser could not sniff it (notably for
  // .md, .csv and source files served without an extension mapping). Fall
  // back to `application/octet-stream`; the extractor then dispatches on the
  // filename extension instead, which is the more reliable signal anyway.
  const mimeType = file.type || 'application/octet-stream'
  const dir = path.join(ATTACHMENTS_DIR, id)
  fs.mkdirSync(dir, { recursive: true })
  const filePath = path.join(dir, path.basename(file.name) || 'attachment')
  fs.writeFileSync(filePath, file.buffer)
  const meta: AttachmentMeta = {
    id,
    fileName: path.basename(file.name) || 'attachment',
    mimeType,
    kind,
    tokenEstimate: 0,
    path: filePath,
  }
  store.set(id, meta)
  return { id, fileName: meta.fileName, mimeType, kind }
}

export function getMeta(id: string): AttachmentMeta | null {
  return store.get(id) ?? null
}

export function getBytes(id: string): { buffer: Buffer; mimeType: string } | null {
  const meta = store.get(id)
  if (!meta) return null
  try {
    const buffer = fs.readFileSync(meta.path)
    return { buffer, mimeType: meta.mimeType }
  } catch {
    // The file was deleted out from under us (manual cleanup, disk error).
    // Treat as missing rather than throwing — a 404 is the honest answer.
    return null
  }
}

export function deleteAttachment(id: string): boolean {
  const meta = store.get(id)
  if (!meta) return false
  const dir = path.dirname(meta.path)
  fs.rmSync(dir, { recursive: true, force: true })
  return store.delete(id)
}

/** Wipe the entire attachment store — both the in-memory map and the on-disk
 *  directory. Called on host startup so attachments from a crashed or killed
 *  previous run do not accumulate and do not get their ids reused (the pane
 *  has no handle to them by the time a new session starts). */
export function cleanup(): void {
  store.clear()
  fs.rmSync(ATTACHMENTS_DIR, { recursive: true, force: true })
  fs.mkdirSync(ATTACHMENTS_DIR, { recursive: true })
}