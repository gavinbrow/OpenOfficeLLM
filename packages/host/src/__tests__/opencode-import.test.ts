// Regression tests for the opencode importer.
//
// The bugs these cover were all "the import silently returned almost nothing on
// a real install": credentials live in a different directory than the config,
// a provider entry legitimately has no baseURL, and the model *ids* are the
// record keys rather than the display names beside them.

import { describe, it, expect, beforeEach, afterEach } from 'vitest'
import fs from 'node:fs'
import path from 'node:path'
import os from 'node:os'
import { importFromOpencode, stripJsonComments } from '../opencode-import.js'

const TMP_DIR = path.join(os.tmpdir(), `ool-opencode-${process.pid}-${Date.now()}`)
const CONFIG_DIR = path.join(TMP_DIR, 'config')
const DATA_DIR = path.join(TMP_DIR, 'data', 'opencode')

const saved: Record<string, string | undefined> = {}
const ENV_KEYS = [
  'OPENCODE_CONFIG_DIR',
  'OPENCODE_DATA_DIR',
  'XDG_DATA_HOME',
  'XDG_CONFIG_HOME',
  'APPDATA',
  'LOCALAPPDATA',
]

function writeConfig(body: unknown, name = 'opencode.jsonc'): void {
  fs.mkdirSync(CONFIG_DIR, { recursive: true })
  fs.writeFileSync(path.join(CONFIG_DIR, name), JSON.stringify(body, null, 2))
}

function writeAuth(body: unknown): void {
  fs.mkdirSync(DATA_DIR, { recursive: true })
  fs.writeFileSync(path.join(DATA_DIR, 'auth.json'), JSON.stringify(body, null, 2))
}

describe('opencode import', () => {
  beforeEach(() => {
    fs.rmSync(TMP_DIR, { recursive: true, force: true })
    for (const k of ENV_KEYS) saved[k] = process.env[k]
    // Both overrides are exclusive, so a real opencode install on the machine
    // running the tests cannot leak into the fixture.
    process.env.OPENCODE_CONFIG_DIR = CONFIG_DIR
    process.env.OPENCODE_DATA_DIR = DATA_DIR
    delete process.env.XDG_DATA_HOME
    delete process.env.XDG_CONFIG_HOME
    delete process.env.APPDATA
    delete process.env.LOCALAPPDATA
  })

  afterEach(() => {
    for (const k of ENV_KEYS) {
      if (saved[k] === undefined) delete process.env[k]
      else process.env[k] = saved[k]
    }
    fs.rmSync(TMP_DIR, { recursive: true, force: true })
  })

  it('reports where it looked when nothing is installed', () => {
    const r = importFromOpencode(false)
    expect(r.ok).toBe(false)
    expect(r.searched).toContain(CONFIG_DIR)
    expect(r.errors[0]).toMatch(/No opencode config/i)
  })

  it('imports a provider that spells out its own baseURL', () => {
    writeConfig({
      provider: {
        lmstudioserver: {
          name: 'LM-Studio',
          options: { baseURL: 'https://lms.example.com/v1', apiKey: 'sk-x' },
          models: { 'qwen3.5-122b': { name: 'Qwen 3.5' }, 'qwen/qwen3.6-27b': {} },
        },
      },
    })
    const r = importFromOpencode(false)
    expect(r.ok).toBe(true)
    expect(r.providers).toHaveLength(1)
    expect(r.providers[0]).toMatchObject({
      id: 'lmstudioserver',
      name: 'LM-Studio',
      baseUrl: 'https://lms.example.com/v1',
    })
    // Model ids come from the keys — `name` is a display label, and sending it
    // upstream as the model id is a 404.
    expect(r.providers[0].models).toEqual(['qwen3.5-122b', 'qwen/qwen3.6-27b'])
  })

  it('fills in a known provider that has no baseURL in the config', () => {
    // Exactly the shape `opencode auth login` leaves behind for Ollama Cloud:
    // a config entry with model overrides only, and the key in auth.json.
    writeConfig({ provider: { 'ollama-cloud': { models: { 'kimi-k3': {} } } } })
    writeAuth({ 'ollama-cloud': { type: 'api', key: 'secret' } })
    const r = importFromOpencode(false)
    expect(r.errors).toEqual([])
    // ollama-cloud has a first-class adapter here, so it is linked rather than
    // re-created as a generic OpenAI-compatible provider.
    expect(r.linkedProviders).toContain('ollama-cloud')
    expect(r.providers.find((p) => p.id === 'ollama-cloud')).toBeUndefined()
  })

  it('finds a provider present only in auth.json', () => {
    writeConfig({})
    writeAuth({ anthropic: { type: 'api', key: 'sk-ant-x' } })
    const r = importFromOpencode(false)
    expect(r.linkedProviders).toContain('anthropic')
  })

  it('skips oauth credentials with an explanation rather than silently', () => {
    writeConfig({})
    writeAuth({ 'github-copilot': { type: 'oauth', key: '' } })
    const r = importFromOpencode(false)
    expect(r.linkedProviders).not.toContain('github-copilot')
    expect(r.errors.join(' ')).toMatch(/github-copilot/)
  })

  it('imports MCP servers disabled, preserving headers', () => {
    writeConfig({
      mcp: {
        chemistry: {
          type: 'remote',
          url: 'https://mcp.example.com/mcp',
          enabled: true,
          headers: { Authorization: 'Bearer x' },
        },
        local: { type: 'local', command: ['npx', '-y', 'server'] },
      },
    })
    const r = importFromOpencode(false)
    expect(r.mcpServers).toHaveLength(2)
    const http = r.mcpServers.find((s) => s.id === 'chemistry')!
    expect(http.transport).toBe('http')
    expect(http.headers).toEqual({ Authorization: 'Bearer x' })
    // Imported servers are always off: consent is the user's to give.
    expect(r.mcpServers.every((s) => s.enabled === false)).toBe(true)
    // `command` given as a whole argv array must split into command + args.
    const stdio = r.mcpServers.find((s) => s.id === 'local')!
    expect(stdio.command).toBe('npx')
    expect(stdio.args).toEqual(['-y', 'server'])
  })

  it('imports agents without treating opencode’s agent mode as an edit mode', () => {
    writeConfig({
      agent: {
        'kimi-coder': {
          description: 'A subagent',
          mode: 'subagent',
          model: 'ollama-cloud/kimi-k2.7-code',
          permission: { bash: 'allow', edit: { 'src/**': 'allow' } },
        },
      },
    })
    const r = importFromOpencode(false)
    expect(r.agents).toHaveLength(1)
    expect(r.agents[0].mode).toBeUndefined()
    // "subagent" is opencode routing, not one of our edit modes — no warning.
    expect(r.errors).toEqual([])
    expect(r.agents[0].enabled).toBe(false)
    // Nested permission objects flatten so the config schema accepts them.
    expect(r.agents[0].permissions).toEqual({ bash: 'allow', edit: 'conditional' })
  })

  it('honours disabled_providers', () => {
    writeConfig({
      disabled_providers: ['lmstudioserver'],
      provider: { lmstudioserver: { options: { baseURL: 'https://x/v1' } } },
    })
    const r = importFromOpencode(false)
    expect(r.providers).toHaveLength(0)
  })
})

describe('stripJsonComments', () => {
  it('removes line and block comments and trailing commas', () => {
    const text = `{
      // a line comment
      "a": 1, /* block */
      "b": "https://example.com//not-a-comment",
    }`
    expect(JSON.parse(stripJsonComments(text))).toEqual({
      a: 1,
      b: 'https://example.com//not-a-comment',
    })
  })
})
