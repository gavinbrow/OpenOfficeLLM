// Minimal YAML-frontmatter parser for skill files.
//
// Deliberately not a YAML library. The skill schema is flat — scalars, string
// lists, booleans — and pulling in a full YAML parser to read six known keys
// would add a parser with a real CVE history to a component that reads files
// from a user-writable directory. Anything this parser does not understand is
// ignored rather than guessed at, and the loader reports the file as invalid
// instead of loading a half-parsed skill.

export interface Frontmatter {
  data: Record<string, string | string[] | boolean>
  body: string
}

const DELIM = /^---[ \t]*\r?$|^---[ \t]*$/

/** Strip one layer of matching quotes, if present. */
function unquote(value: string): string {
  const v = value.trim()
  if (v.length >= 2 && ((v[0] === '"' && v.endsWith('"')) || (v[0] === "'" && v.endsWith("'")))) {
    return v.slice(1, -1)
  }
  return v
}

function coerce(raw: string): string | string[] | boolean {
  const v = raw.trim()
  if (v === 'true') return true
  if (v === 'false') return false
  // Inline list: [a, b, c]
  if (v.startsWith('[') && v.endsWith(']')) {
    const inner = v.slice(1, -1).trim()
    if (!inner) return []
    return inner
      .split(',')
      .map((s) => unquote(s))
      .filter((s) => s.length > 0)
  }
  return unquote(v)
}

/**
 * Split a markdown file into frontmatter data and body.
 *
 * A file with no leading `---` is not an error: it is a skill whose entire
 * content is the prompt, named after its filename by the loader.
 */
export function parseFrontmatter(text: string): Frontmatter {
  // Strip a BOM — a skill authored in Notepad has one, and it would otherwise
  // make the opening `---` fail to match and silently disable the frontmatter.
  const source = text.charCodeAt(0) === 0xfeff ? text.slice(1) : text
  const lines = source.split('\n')
  if (lines.length === 0 || !DELIM.test(lines[0].trim())) {
    return { data: {}, body: source }
  }

  let end = -1
  for (let i = 1; i < lines.length; i++) {
    if (DELIM.test(lines[i].trim())) {
      end = i
      break
    }
  }
  if (end === -1) return { data: {}, body: source }

  const data: Record<string, string | string[] | boolean> = {}
  let pendingKey: string | null = null
  let pendingList: string[] = []

  const commitList = () => {
    if (pendingKey !== null) {
      data[pendingKey] = pendingList
      pendingKey = null
      pendingList = []
    }
  }

  for (let i = 1; i < end; i++) {
    const line = lines[i].replace(/\r$/, '')
    if (!line.trim() || line.trim().startsWith('#')) continue

    // Block list item: "  - value"
    const listItem = /^\s*-\s+(.*)$/.exec(line)
    if (listItem && pendingKey !== null) {
      pendingList.push(unquote(listItem[1]))
      continue
    }

    const colon = line.indexOf(':')
    if (colon === -1) continue
    commitList()
    const key = line.slice(0, colon).trim()
    const value = line.slice(colon + 1)
    if (!key) continue
    if (value.trim() === '') {
      // Either an empty scalar or the header of a block list; decided by
      // whether the next non-blank line is a list item.
      pendingKey = key
      pendingList = []
      continue
    }
    data[key] = coerce(value)
  }
  commitList()

  return {
    data,
    body: lines
      .slice(end + 1)
      .join('\n')
      .trim(),
  }
}
