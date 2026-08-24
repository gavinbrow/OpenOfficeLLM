// Which tabs the side panel will try to read.
//
// Getting this wrong is not cosmetic: attempting to inject into a chrome://
// page throws an opaque error that surfaces to the user as "the assistant is
// broken", when the truthful answer is "no extension may read this page".

import { describe, it, expect } from 'vitest'
import { isScriptableUrl } from '../bootstrap'

describe('isScriptableUrl', () => {
  it('accepts ordinary web pages', () => {
    expect(isScriptableUrl('https://example.com/article')).toBe(true)
    expect(isScriptableUrl('http://localhost:3000/')).toBe(true)
  })

  it('accepts file URLs, leaving the real check to Chrome', () => {
    // Whether file access is allowed is a per-install toggle no API exposes,
    // so guessing "no" here would refuse a case that often works.
    expect(isScriptableUrl('file:///C:/notes.html')).toBe(true)
  })

  it('refuses browser-internal pages no extension may script', () => {
    expect(isScriptableUrl('chrome://settings')).toBe(false)
    expect(isScriptableUrl('chrome-extension://abc/page.html')).toBe(false)
    expect(isScriptableUrl('edge://flags')).toBe(false)
    expect(isScriptableUrl('about:blank')).toBe(false)
    expect(isScriptableUrl('devtools://devtools/bundled/inspector.html')).toBe(false)
    expect(isScriptableUrl('view-source:https://example.com')).toBe(false)
  })

  it('refuses the Chrome Web Store, which is blocked by policy', () => {
    expect(isScriptableUrl('https://chromewebstore.google.com/detail/x')).toBe(false)
    expect(isScriptableUrl('https://chrome.google.com/webstore/category/extensions')).toBe(false)
  })

  it('refuses schemes with no page to read', () => {
    expect(isScriptableUrl('ftp://files.example.com')).toBe(false)
    expect(isScriptableUrl('data:text/html,hi')).toBe(false)
    expect(isScriptableUrl('javascript:void(0)')).toBe(false)
  })

  it('refuses empty and unparseable input rather than throwing', () => {
    expect(isScriptableUrl('')).toBe(false)
    expect(isScriptableUrl('not a url')).toBe(false)
  })
})
