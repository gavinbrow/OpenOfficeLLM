// Flatten message content to a plain string. The pane only ever produces
// string content — image content blocks are created host-side in the prompt
// builder — so every UI display path is showing text. This helper narrows the
// `string | ContentBlock[]` union defensively; image blocks contribute no
// text (their payload is image bytes the host holds, not displayable text).

import type { ContentBlock } from '@openofficellm/shared'

export function textOf(content: string | ContentBlock[]): string {
  if (typeof content === 'string') return content
  return content
    .filter((b): b is Extract<ContentBlock, { type: 'text' }> => b.type === 'text')
    .map((b) => b.text)
    .join('')
}