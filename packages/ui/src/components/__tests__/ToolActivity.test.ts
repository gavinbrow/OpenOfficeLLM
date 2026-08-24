import { describe, it, expect } from 'vitest'
import type { ChatMessage } from '@openofficellm/shared'
import { buildTranscript, summarizeArgs, type TranscriptRow } from '../ToolActivity'

const OPTS = { showReasoning: true }

function user(id: string, content: string): ChatMessage {
  return { id, role: 'user', content }
}

function assistant(
  id: string,
  content: string,
  toolCalls?: [string, string, string][],
): ChatMessage {
  return {
    id,
    role: 'assistant',
    content,
    ...(toolCalls
      ? { toolCalls: toolCalls.map(([tid, name, args]) => ({ id: tid, name, arguments: args })) }
      : {}),
  }
}

function toolResult(id: string, callId: string, name: string, content: string): ChatMessage {
  return { id, role: 'tool', content, toolCallId: callId, toolName: name }
}

function activityAt(rows: TranscriptRow[], i: number) {
  const row = rows[i]
  if (row.kind !== 'activity') throw new Error(`row ${i} is a ${row.kind}, not an activity`)
  return row.group
}

describe('buildTranscript', () => {
  it('leaves a plain exchange untouched', () => {
    const rows = buildTranscript([user('u1', 'hi'), assistant('a1', 'hello')], OPTS)
    expect(rows.map((r) => r.kind)).toEqual(['message', 'message'])
  })

  it('folds a whole tool run into one group between question and answer', () => {
    const rows = buildTranscript(
      [
        user('u1', 'bold the headings'),
        assistant('a1', '', [['c1', 'read_document', '{}']]),
        toolResult('t1', 'c1', 'read_document', 'Heading one\nBody'),
        assistant('a2', '', [['c2', 'format_text', '{"bold":true}']]),
        toolResult('t2', 'c2', 'format_text', 'formatted paragraph 1'),
        assistant('a3', 'Done — the headings are bold.'),
      ],
      OPTS,
    )

    expect(rows.map((r) => r.kind)).toEqual(['message', 'activity', 'message'])
    const group = activityAt(rows, 1)
    expect(group.steps.map((s) => s.name)).toEqual(['read_document', 'format_text'])
    expect(group.steps[0].result).toBe('Heading one\nBody')
    expect(group.steps[1].result).toBe('formatted paragraph 1')
  })

  it('keeps assistant prose that accompanies a tool call, and starts a group after it', () => {
    const rows = buildTranscript(
      [
        user('u1', 'check it'),
        assistant('a1', 'Let me look.', [['c1', 'read_document', '{}']]),
        toolResult('t1', 'c1', 'read_document', 'text'),
        assistant('a2', 'Looks fine.'),
      ],
      OPTS,
    )
    expect(rows.map((r) => r.kind)).toEqual(['message', 'message', 'activity', 'message'])
    expect(activityAt(rows, 2).steps).toHaveLength(1)
  })

  it('splits groups around an intervening answer rather than merging them', () => {
    const rows = buildTranscript(
      [
        assistant('a1', '', [['c1', 'read_document', '{}']]),
        toolResult('t1', 'c1', 'read_document', 'x'),
        assistant('a2', 'Halfway there.'),
        assistant('a3', '', [['c2', 'format_text', '{}']]),
        toolResult('t2', 'c2', 'format_text', 'y'),
      ],
      OPTS,
    )
    expect(rows.map((r) => r.kind)).toEqual(['activity', 'message', 'activity'])
    expect(activityAt(rows, 0).steps).toHaveLength(1)
    expect(activityAt(rows, 2).steps).toHaveLength(1)
  })

  it('marks a call with no result yet as still running', () => {
    const rows = buildTranscript([assistant('a1', '', [['c1', 'insert_table', '{}']])], OPTS)
    expect(activityAt(rows, 0).steps[0].result).toBeUndefined()
  })

  it('shows a result whose call never arrived instead of dropping it', () => {
    const rows = buildTranscript([toolResult('t1', 'missing', 'replace_all', 'ok')], OPTS)
    const step = activityAt(rows, 0).steps[0]
    expect(step.name).toBe('replace_all')
    expect(step.result).toBe('ok')
  })

  it('pairs repeated calls to the same tool with their own results in order', () => {
    const rows = buildTranscript(
      [
        assistant('a1', '', [
          ['c1', 'format_text', '{"paragraph":1}'],
          ['c2', 'format_text', '{"paragraph":2}'],
        ]),
        toolResult('t1', 'c1', 'format_text', 'first'),
        toolResult('t2', 'c2', 'format_text', 'second'),
      ],
      OPTS,
    )
    expect(activityAt(rows, 0).steps.map((s) => s.result)).toEqual(['first', 'second'])
  })

  it('renders the streaming placeholder only until the first tool call lands', () => {
    const pending = assistant('a1', '')
    expect(buildTranscript([pending], { ...OPTS, streamingId: 'a1' }).map((r) => r.kind)).toEqual([
      'message',
    ])

    const calling = assistant('a1', '', [['c1', 'read_document', '{}']])
    expect(buildTranscript([calling], { ...OPTS, streamingId: 'a1' }).map((r) => r.kind)).toEqual([
      'activity',
    ])
  })

  it('keeps a reasoning-only message visible, and hides it when reasoning is off', () => {
    const msg: ChatMessage = { id: 'a1', role: 'assistant', content: '', reasoning: 'hmm' }
    expect(buildTranscript([msg], { showReasoning: true }).map((r) => r.kind)).toEqual(['message'])
    expect(buildTranscript([msg], { showReasoning: false })).toEqual([])
  })
})

describe('summarizeArgs', () => {
  it('drops the nulls models emit for unused parameters', () => {
    expect(summarizeArgs('{"paragraph":null,"find":"policy","from":null}')).toBe('find=policy')
  })

  it('truncates a long value rather than pasting the whole replacement text', () => {
    const out = summarizeArgs(JSON.stringify({ text: 'x'.repeat(200) }))
    expect(out.length).toBeLessThan(60)
    expect(out).toContain('…')
  })

  it('falls back to the raw string when the arguments are not valid JSON', () => {
    expect(summarizeArgs('not json')).toBe('not json')
    expect(summarizeArgs('')).toBe('')
  })
})
