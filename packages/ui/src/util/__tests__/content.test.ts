import { describe, it, expect } from 'vitest'
import { textOf } from '../content'

describe('textOf', () => {
  it('returns a plain string unchanged', () => {
    expect(textOf('hello')).toBe('hello')
  })

  it('joins the text of text content blocks', () => {
    expect(textOf([{ type: 'text', text: 'hi' }])).toBe('hi')
  })

  it('drops image blocks and keeps only text blocks', () => {
    const out = textOf([
      { type: 'image', mimeType: 'image/png', data: 'AAA' },
      { type: 'text', text: 'caption' },
    ])
    expect(out).toBe('caption')
  })

  it('returns an empty string for an empty block array', () => {
    expect(textOf([])).toBe('')
  })

  it('returns an empty string when only image blocks are present', () => {
    expect(textOf([{ type: 'image', mimeType: 'image/png', data: 'AAA' }])).toBe('')
  })

  it('joins multiple text blocks without a separator', () => {
    expect(
      textOf([
        { type: 'text', text: 'a' },
        { type: 'text', text: 'b' },
        { type: 'text', text: 'c' },
      ]),
    ).toBe('abc')
  })
})