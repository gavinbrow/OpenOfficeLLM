// Loads skills from every source and merges them by id (P5.2).
//
// Precedence, lowest to highest: built-in → opencode → user. A user skill named
// `proofread.md` replaces the built-in of that id; the built-in is not merged
// field-by-field, because a half-overridden prompt is harder to reason about
// than a replaced one.
//
// The opencode source is strictly READ-ONLY. We never write to opencode's
// directories, and a malformed file there is skipped with a warning rather than
// treated as a failure of the whole load.

import fs from 'node:fs'
import path from 'node:path'
import os from 'node:os'
import type { ContextScope, DetectedHost, EditMode, HostKind, Skill } from '@openofficellm/shared'
import { resolveSkillsDir } from '../paths.js'
import { logger } from '../logging.js'
import { parseFrontmatter } from './frontmatter.js'
import { builtinSkills } from './builtin.js'

const VALID_MODES: EditMode[] = ['propose', 'direct', 'agentic']
const VALID_SCOPES: ContextScope[] = [
  'none',
  'selection',
  'paragraph',
  'document',
  'sheet',
  'range',
]

function asString(v: unknown): string | undefined {
  return typeof v === 'string' && v.trim() ? v.trim() : undefined
}

const HOST_KINDS: readonly HostKind[] = ['word', 'excel', 'browser']

function asHosts(v: unknown): HostKind[] {
  const list = Array.isArray(v) ? v : typeof v === 'string' ? v.split(/[,\s]+/) : []
  const out = list
    .map((h) => String(h).trim().toLowerCase())
    .filter((h): h is HostKind => (HOST_KINDS as readonly string[]).includes(h))
  return [...new Set(out)]
}

/** Derive an id from a filename: `Rewrite Formally.md` → `rewrite-formally`. */
function idFromFilename(file: string): string {
  return path
    .basename(file, path.extname(file))
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, '-')
    .replace(/^-+|-+$/g, '')
}

export function parseSkillFile(
  text: string,
  file: string,
  source: 'user' | 'opencode',
): Skill | null {
  const { data, body } = parseFrontmatter(text)
  const prompt = body.trim()
  const id = asString(data.id) ?? idFromFilename(file)
  if (!id) return null
  // A skill with no prompt cannot do anything; loading it would put a dead
  // button in the skill bar.
  if (!prompt) return null

  const mode = asString(data.mode)
  const scope = asString(data.contextScope) ?? asString(data.scope)

  return {
    id,
    name: asString(data.name) ?? asString(data.title) ?? id,
    description: asString(data.description) ?? '',
    hosts: asHosts(data.hosts),
    mode: mode && VALID_MODES.includes(mode as EditMode) ? (mode as EditMode) : undefined,
    model: asString(data.model),
    icon: asString(data.icon),
    contextScope:
      scope && VALID_SCOPES.includes(scope as ContextScope) ? (scope as ContextScope) : undefined,
    prompt,
    builtIn: false,
    source,
    // Only user skills get a path: it is what the editor writes back to, and
    // handing out an opencode path would invite exactly the write we promised
    // never to make.
    path: source === 'user' ? file : undefined,
  }
}

function loadDir(dir: string, source: 'user' | 'opencode'): Skill[] {
  let entries: fs.Dirent[]
  try {
    entries = fs.readdirSync(dir, { withFileTypes: true })
  } catch {
    return []
  }
  const out: Skill[] = []
  for (const entry of entries) {
    let file: string | null = null
    if (entry.isFile() && /\.(md|markdown)$/i.test(entry.name)) {
      file = path.join(dir, entry.name)
    } else if (entry.isDirectory()) {
      // opencode packages a skill as `<name>/SKILL.md`.
      const nested = path.join(dir, entry.name, 'SKILL.md')
      if (fs.existsSync(nested)) file = nested
    }
    if (!file) continue
    try {
      const text = fs.readFileSync(file, 'utf8')
      const skill = parseSkillFile(text, file, source)
      if (skill) out.push(skill)
      else logger.warn({ msg: 'skipping skill with no prompt body', file })
    } catch (e) {
      logger.warn({ msg: 'failed to read skill', file, error: String((e as Error).message ?? e) })
    }
  }
  return out
}

/** opencode's skills live beside its config. Read-only. */
function opencodeSkillDirs(): string[] {
  const dirs: string[] = []
  const explicit = process.env.OPENCODE_CONFIG_DIR
  if (explicit) dirs.push(path.join(explicit, 'skills'))
  dirs.push(path.join(os.homedir(), '.config', 'opencode', 'skills'))
  return dirs
}

export interface LoadedSkills {
  skills: Skill[]
  /** Directories actually scanned, for the settings UI to show the user where
   *  to drop a file. */
  userDir: string
}

export function loadSkills(): LoadedSkills {
  const userDir = resolveSkillsDir()
  const byId = new Map<string, Skill>()

  for (const s of builtinSkills()) byId.set(s.id, s)
  for (const dir of opencodeSkillDirs()) {
    for (const s of loadDir(dir, 'opencode')) byId.set(s.id, s)
  }
  for (const s of loadDir(userDir, 'user')) byId.set(s.id, s)

  const skills = [...byId.values()].sort((a, b) => a.name.localeCompare(b.name))
  return { skills, userDir }
}

/** Skills applicable to a host. A skill with no `hosts` applies to all of them. */
export function skillsForHost(skills: Skill[], host: DetectedHost): Skill[] {
  if (host === 'none') return skills
  return skills.filter((s) => s.hosts.length === 0 || s.hosts.includes(host))
}

export function findSkill(skills: Skill[], id: string): Skill | undefined {
  return skills.find((s) => s.id === id)
}

// ─── Cache ───────────────────────────────────────────────────────────────
//
// Skills are read on every /api/skills request in development, which makes
// editing a skill file a save-and-refresh loop. The cache exists only to keep
// a chat turn (which resolves a skill prompt) from hitting the disk.

let cache: LoadedSkills | null = null
let cachedAt = 0
const CACHE_TTL_MS = 2_000

export function getSkills(force = false): LoadedSkills {
  const now = Date.now()
  if (!force && cache && now - cachedAt < CACHE_TTL_MS) return cache
  cache = loadSkills()
  cachedAt = now
  return cache
}

export function bustSkillCache(): void {
  cache = null
  cachedAt = 0
}

/** Write a user skill to disk. Returns the path written. */
export function saveUserSkill(skill: Skill): string {
  const dir = resolveSkillsDir()
  fs.mkdirSync(dir, { recursive: true })
  // Derive the filename from the id, never from user-supplied path input —
  // this endpoint is reachable from the webview, and a `name` of `..\..\evil`
  // would otherwise write outside the skills directory.
  const safeId = skill.id
    .toLowerCase()
    .replace(/[^a-z0-9-]+/g, '-')
    .replace(/^-+|-+$/g, '')
  if (!safeId) throw new Error('invalid skill id')
  const file = path.join(dir, `${safeId}.md`)

  const fm: string[] = ['---']
  fm.push(`id: ${safeId}`)
  fm.push(`name: ${JSON.stringify(skill.name)}`)
  if (skill.description) fm.push(`description: ${JSON.stringify(skill.description)}`)
  if (skill.hosts.length > 0) fm.push(`hosts: [${skill.hosts.join(', ')}]`)
  if (skill.mode) fm.push(`mode: ${skill.mode}`)
  if (skill.model) fm.push(`model: ${JSON.stringify(skill.model)}`)
  if (skill.icon) fm.push(`icon: ${JSON.stringify(skill.icon)}`)
  if (skill.contextScope) fm.push(`contextScope: ${skill.contextScope}`)
  fm.push('---', '', skill.prompt ?? '')

  fs.writeFileSync(file, fm.join('\n'), 'utf8')
  bustSkillCache()
  return file
}

/** Delete a user skill. Built-ins and opencode skills are not deletable. */
export function deleteUserSkill(id: string): boolean {
  const safeId = id
    .toLowerCase()
    .replace(/[^a-z0-9-]+/g, '-')
    .replace(/^-+|-+$/g, '')
  if (!safeId) return false
  const file = path.join(resolveSkillsDir(), `${safeId}.md`)
  try {
    fs.unlinkSync(file)
    bustSkillCache()
    return true
  } catch {
    return false
  }
}
