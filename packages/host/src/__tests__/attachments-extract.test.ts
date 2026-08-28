// Text extraction dispatch.
//
// The PDF/DOCX/XLSX/PPTX extractors each need a real binary of their format
// and a working dynamic import of a heavy parser, so they are covered by
// integration tests rather than this unit suite. This file covers the
// in-memory branches: utf-8 text reads, the image short-circuit, the
// unsupported-type error, and truncation at the 48k character cap.

import { describe, it, expect } from 'vitest'
import { extractText, AttachmentError } from '../attachments/extract.js'

describe('extractText', () => {
  it('reads a .txt file as utf-8 text', async () => {
    const out = await extractText(Buffer.from('hello world', 'utf-8'), 'notes.txt', 'text/plain')
    expect(out.kind).toBe('text')
    expect(out.text).toBe('hello world')
    expect(out.truncated).toBe(false)
  })

  it('reads a .md file by extension even with an empty mime type', async () => {
    const out = await extractText(Buffer.from('# heading', 'utf-8'), 'README.md', '')
    expect(out.kind).toBe('text')
    expect(out.text).toBe('# heading')
  })

  it('reads a text/* mime type regardless of extension', async () => {
    const out = await extractText(Buffer.from('body', 'utf-8'), 'weird.ext', 'text/plain')
    expect(out.text).toBe('body')
    expect(out.kind).toBe('text')
  })

  it('returns kind image with empty text for image/* mime types', async () => {
    const out = await extractText(
      Buffer.from([0x89, 0x50, 0x4e, 0x47]),
      'photo.png',
      'image/png',
    )
    expect(out.kind).toBe('image')
    expect(out.text).toBe('')
    expect(out.truncated).toBe(false)
  })

  it('throws AttachmentError with code unsupported_file_type for an unknown extension', async () => {
    await expect(
      extractText(Buffer.from([0x00, 0x01]), 'malware.exe', 'application/octet-stream'),
    ).rejects.toMatchObject({
      name: 'AttachmentError',
      code: 'unsupported_file_type',
    })
  })

  it('throws AttachmentError for an extensionless file with an unknown basename', async () => {
    await expect(
      extractText(Buffer.from([0x00]), 'binaryblob', 'application/octet-stream'),
    ).rejects.toBeInstanceOf(AttachmentError)
  })

  it('truncates text longer than 48,000 characters and marks it', async () => {
    const long = 'x'.repeat(50_000)
    const out = await extractText(Buffer.from(long, 'utf-8'), 'big.txt', 'text/plain')
    expect(out.truncated).toBe(true)
    expect(out.text).toContain('[truncated]')
    // The kept prefix plus the marker is shorter than the original.
    expect(out.text.length).toBeLessThan(long.length)
    // The first 48,000 characters are preserved verbatim.
    expect(out.text.startsWith('x'.repeat(48_000))).toBe(true)
  })

  it('leaves text under the cap untruncated', async () => {
    const body = 'y'.repeat(48_000)
    const out = await extractText(Buffer.from(body, 'utf-8'), 'exact.txt', 'text/plain')
    expect(out.truncated).toBe(false)
    expect(out.text).toBe(body)
  })
})

// PDF (.pdf), Word (.docx), Excel (.xlsx) and PowerPoint (.pptx) extraction
// are integration-tested: each needs a valid binary of its format and a
// working dynamic import of pdfjs-dist / mammoth / exceljs / jszip. Those
// suites live alongside the format fixtures and are not unit-tested here.