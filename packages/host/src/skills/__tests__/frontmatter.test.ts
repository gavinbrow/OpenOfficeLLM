import { describe, it, expect } from 'vitest'
import path from 'node:path'
import { parseFrontmatter } from '../frontmatter.js'
import { parseSkillFile } from '../loader.js'

describe('parseFrontmatter', () => {
  it('returns the whole file as body when there is no frontmatter', () => {
    const r = parseFrontmatter('Just a prompt.')
    expect(r.data).toEqual({})
    expect(r.body).toBe('Just a prompt.')
  })

  it('parses scalars, booleans, and inline lists', () => {
    const r = parseFrontmatter(
      ['---', 'name: Proofread', 'builtIn: false', 'hosts: [word, excel]', '---', 'Body'].join(
        '\n',
      ),
    )
    expect(r.data.name).toBe('Proofread')
    expect(r.data.builtIn).toBe(false)
    expect(r.data.hosts).toEqual(['word', 'excel'])
    expect(r.body).toBe('Body')
  })

  it('parses block lists', () => {
    const r = parseFrontmatter(['---', 'hosts:', '  - word', '  - excel', '---', 'Body'].join('\n'))
    expect(r.data.hosts).toEqual(['word', 'excel'])
  })

  it('strips quotes from values', () => {
    const r = parseFrontmatter(['---', 'name: "With: a colon"', '---', 'x'].join('\n'))
    expect(r.data.name).toBe('With: a colon')
  })

  // A file saved from Notepad starts with a BOM, which would otherwise make the
  // opening delimiter fail to match and silently discard all frontmatter.
  it('tolerates a UTF-8 BOM', () => {
    const r = parseFrontmatter(['\ufeff---', 'name: X', '---', 'Body'].join('\n'))
    expect(r.data.name).toBe('X')
    expect(r.body).toBe('Body')
  })

  it('treats an unterminated frontmatter block as body text', () => {
    const r = parseFrontmatter(['---', 'name: X', 'no closing delimiter'].join('\n'))
    expect(r.data).toEqual({})
  })

  it('ignores comments and blank lines', () => {
    const r = parseFrontmatter(['---', '# a comment', '', 'name: X', '---', 'b'].join('\n'))
    expect(r.data).toEqual({ name: 'X' })
  })
})

describe('parseSkillFile', () => {
  it('derives id and name from the filename when absent', () => {
    // path.join so the fixture uses the native separator on both platforms —
    // idFromFilename delegates to path.basename, which is separator-aware.
    const skill = parseSkillFile(
      'Do the thing.',
      path.join('C:', 'skills', 'Rewrite Formally.md'),
      'user',
    )
    expect(skill?.id).toBe('rewrite-formally')
    expect(skill?.name).toBe('rewrite-formally')
    expect(skill?.prompt).toBe('Do the thing.')
  })

  it('rejects a skill with no prompt body', () => {
    // Loading one would put a button in the skill bar that does nothing.
    expect(
      parseSkillFile(['---', 'name: Empty', '---', ''].join('\n'), 'empty.md', 'user'),
    ).toBeNull()
  })

  it('ignores an invalid mode rather than failing the file', () => {
    const skill = parseSkillFile(
      ['---', 'name: X', 'mode: sideways', '---', 'body'].join('\n'),
      'x.md',
      'user',
    )
    expect(skill?.mode).toBeUndefined()
  })

  // opencode's directories are read-only to us; handing out a path would invite
  // an editor write straight into them.
  it('records a path for user skills only', () => {
    expect(parseSkillFile('body', 'a.md', 'user')?.path).toBe('a.md')
    expect(parseSkillFile('body', 'a.md', 'opencode')?.path).toBeUndefined()
  })
})
