// The built-in skill pack (P5.3).
//
// These are shipped as data rather than as files on disk so a fresh install has
// working skills before the user has ever opened the skills folder. A user
// skill with the same id overrides the built-in — see the loader's merge order.

import type { Skill } from '@openofficellm/shared'

function skill(s: Omit<Skill, 'builtIn' | 'source'>): Skill {
  return { ...s, builtIn: true, source: 'built-in' }
}

const WORD: Skill[] = [
  skill({
    id: 'rewrite-formally',
    name: 'Rewrite formally',
    description: 'Rewrite the selection in a more formal register.',
    hosts: ['word'],
    contextScope: 'selection',
    icon: '✒️',
    prompt: [
      'Rewrite the selected text in a more formal, professional register.',
      'Preserve the meaning, the level of detail, and any specific figures or names exactly.',
      'Do not add new claims. Do not pad the length — formal does not mean longer.',
      'Return only the rewritten text via the edit tool; do not explain the changes unless asked.',
    ].join(' '),
  }),
  skill({
    id: 'summarize-bullets',
    name: 'Summarize to bullets',
    description: 'Condense the selection or document into bullet points.',
    hosts: ['word'],
    contextScope: 'selection',
    icon: '•',
    prompt: [
      'Summarize the provided text as a tight bulleted list.',
      'One idea per bullet, no sub-bullets unless the structure genuinely nests.',
      'Keep every specific number, date, and name that carries meaning.',
      'Aim for the shortest list that loses nothing important.',
    ].join(' '),
  }),
  skill({
    id: 'tighten',
    name: 'Tighten',
    description: 'Cut wordiness without changing meaning.',
    hosts: ['word'],
    contextScope: 'selection',
    icon: '✂️',
    prompt: [
      'Tighten the selected text. Remove redundancy, hedging, and filler.',
      "Keep the author's voice and every substantive claim.",
      'Do not change terminology, and do not merge paragraphs that make distinct points.',
    ].join(' '),
  }),
  skill({
    id: 'explain-this',
    name: 'Explain this',
    description: 'Explain the selected passage in plain language.',
    hosts: ['word'],
    contextScope: 'selection',
    icon: '💡',
    prompt: [
      'Explain the selected passage in plain language.',
      'Define any jargon on first use. Say what it means, not what it says.',
      'This is an explanation for the reader, not an edit — do not modify the document.',
    ].join(' '),
  }),
  skill({
    id: 'continue-writing',
    name: 'Continue writing',
    description: 'Draft the next passage in the document’s own voice.',
    hosts: ['word'],
    contextScope: 'paragraph',
    icon: '↳',
    prompt: [
      'Continue the document from where it leaves off.',
      'Match the existing voice, tense, formatting conventions, and level of detail.',
      'Write one or two paragraphs — enough to be useful, short enough to redirect.',
      'Insert the continuation after the cursor; do not restate what is already written.',
    ].join(' '),
  }),
  skill({
    id: 'proofread',
    name: 'Proofread',
    description: 'Fix grammar, spelling, and punctuation only.',
    hosts: ['word'],
    contextScope: 'selection',
    icon: '✓',
    prompt: [
      'Proofread the text for grammar, spelling, and punctuation.',
      'Fix only actual errors. Do not restyle, resequence, or "improve" correct prose —',
      'a proofread that rewrites is a rewrite, and the user did not ask for one.',
      'If you find nothing wrong, say so rather than inventing a change.',
    ].join(' '),
  }),
]

const EXCEL: Skill[] = [
  skill({
    id: 'explain-formula',
    name: 'Explain formula',
    description: 'Explain what the selected formula does.',
    hosts: ['excel'],
    contextScope: 'range',
    icon: 'ƒ',
    prompt: [
      'Explain the formula in the selected cell, step by step.',
      'Name each function, say what its arguments refer to, and describe what the',
      'whole expression computes. Flag anything fragile: hardcoded ranges,',
      'missing absolute references, or error-prone lookups.',
      'This is an explanation — do not modify the sheet.',
    ].join(' '),
  }),
  skill({
    id: 'write-formula',
    name: 'Write formula',
    description: 'Build a formula for what the user describes.',
    hosts: ['excel'],
    contextScope: 'range',
    icon: '=',
    prompt: [
      'Write an Excel formula that does what the user asks.',
      'Use the actual column letters and row numbers from the provided context —',
      'never placeholder ranges. Prefer structured, readable formulas over clever',
      'one-liners. Explain the formula in one sentence, then write it to the cell.',
    ].join(' '),
  }),
  skill({
    id: 'clean-data',
    name: 'Clean this data',
    description: 'Normalize inconsistent values in the selected range.',
    hosts: ['excel'],
    contextScope: 'range',
    icon: '🧹',
    prompt: [
      'Identify inconsistencies in the selected range: mixed date formats, stray',
      'whitespace, inconsistent capitalization, numbers stored as text, duplicate',
      'entries that differ only by formatting.',
      'List what you found before changing anything. Never silently drop a row.',
    ].join(' '),
  }),
  skill({
    id: 'summarize-sheet',
    name: 'Summarize sheet',
    description: 'Describe what the sheet contains and how it is structured.',
    hosts: ['excel'],
    contextScope: 'sheet',
    icon: '📋',
    prompt: [
      'Describe what this sheet contains: what each column holds, how many rows,',
      'what the sheet appears to be for, and where the interesting values are.',
      'Remember the rows you can see may be a sample — do not state totals or',
      'counts as fact unless you computed them over the full range.',
    ].join(' '),
  }),
  skill({
    id: 'find-anomalies',
    name: 'Find anomalies',
    description: 'Spot outliers and suspicious values in the data.',
    hosts: ['excel'],
    contextScope: 'range',
    icon: '⚠️',
    prompt: [
      'Look for anomalies: outliers, impossible values, gaps in a sequence,',
      'duplicated keys, totals that do not add up, dates outside a plausible range.',
      'Cite the specific cell for each finding. Say when something is merely',
      'unusual rather than certainly wrong.',
    ].join(' '),
  }),
]

export function builtinSkills(): Skill[] {
  return [...WORD, ...EXCEL]
}
