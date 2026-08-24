import fs from 'node:fs'
import path from 'node:path'
import { resolveOfficejsCacheDir, ensureDirs } from './paths.js'
import { logger } from './logging.js'

const OFFICE_JS_CDN_URL = 'https://appsforoffice.microsoft.com/lib/1/hosted/office.js'
const OFFICE_JS_DEBUG_CDN_URL = 'https://appsforoffice.microsoft.com/lib/1/hosted/office.debug.js'
const CACHE_FILE = 'office.js'
const CACHE_META = 'office.js.meta.json'

const STUB_OFFICE_JS = `
// OpenOfficeLLM offline Office.js stub — CDN was unreachable at cache time.
// Provides a minimal Office.onReady so the task pane can boot and surface a
// recovery message rather than failing on a missing global.
(function () {
  var queue = [];
  var hostInfo = null;
  window.Office = {
    onReady: function (cb) {
      if (typeof cb === 'function') {
        if (hostInfo) { try { cb(hostInfo); } catch (e) {} }
        else { queue.push(cb); }
      }
      return Promise.resolve(hostInfo);
    },
    initialize: function () {},
    context: { host: 'Unknown', platform: 'Unknown', requirements: { isSetSupported: function () { return false; } } },
    select: function () { return { __stub: true }; }
  };
  // Drain once everything else has had a chance to register handlers.
  window.addEventListener('load', function () {
    hostInfo = { host: 'Unknown', platform: 'Unknown' };
    while (queue.length) { try { queue.shift()(hostInfo); } catch (e) {} }
  });
})();
`.trim()

interface CacheMeta {
  fetchedAt: string
  source: 'cdn' | 'debug-cdn' | 'stub'
  bytes: number
}

function cacheDir(): string {
  return resolveOfficejsCacheDir()
}

function writeCache(content: string, source: CacheMeta['source']): void {
  ensureDirs()
  const meta: CacheMeta = {
    fetchedAt: new Date().toISOString(),
    source,
    bytes: Buffer.byteLength(content, 'utf8'),
  }
  fs.writeFileSync(path.join(cacheDir(), CACHE_FILE), content, 'utf8')
  fs.writeFileSync(path.join(cacheDir(), CACHE_META), JSON.stringify(meta, null, 2), 'utf8')
}

async function fetchWithTimeout(url: string, ms: number): Promise<Response> {
  const ctrl = new AbortController()
  const timer = setTimeout(() => ctrl.abort(), ms)
  try {
    return await fetch(url, { signal: ctrl.signal, redirect: 'follow' })
  } finally {
    clearTimeout(timer)
  }
}

export async function refreshOfficeJsCache(): Promise<CacheMeta> {
  for (const item of [
    [OFFICE_JS_CDN_URL, 'cdn'] as const,
    [OFFICE_JS_DEBUG_CDN_URL, 'debug-cdn'] as const,
  ]) {
    const [url, source] = item
    try {
      logger.info({ msg: 'office.js cache: fetching from CDN', url })
      const res = await fetchWithTimeout(url, 5000)
      if (!res.ok) {
        logger.warn({ msg: 'office.js CDN returned non-200', url, status: res.status })
        continue
      }
      const text = await res.text()
      if (!text || text.length < 1000) {
        logger.warn({ msg: 'office.js CDN returned empty body', url })
        continue
      }
      writeCache(text, source)
      logger.info({ msg: 'office.js cache: fetched', source, bytes: text.length })
      return { fetchedAt: new Date().toISOString(), source, bytes: text.length }
    } catch (e) {
      logger.warn({
        msg: 'office.js CDN fetch failed',
        url,
        error: String((e as Error).message ?? e),
      })
    }
  }
  writeCache(STUB_OFFICE_JS, 'stub')
  logger.warn({ msg: 'office.js cache: using offline stub' })
  return { fetchedAt: new Date().toISOString(), source: 'stub', bytes: STUB_OFFICE_JS.length }
}

export function getCachedOfficeJs(): string | null {
  try {
    return fs.readFileSync(path.join(cacheDir(), CACHE_FILE), 'utf8')
  } catch {
    return null
  }
}

export function getOfficeJsMeta(): CacheMeta | null {
  try {
    const raw = fs.readFileSync(path.join(cacheDir(), CACHE_META), 'utf8')
    return JSON.parse(raw) as CacheMeta
  } catch {
    return null
  }
}

export function officeJsCacheExists(): boolean {
  try {
    fs.accessSync(path.join(cacheDir(), CACHE_FILE))
    return true
  } catch {
    return false
  }
}

export async function ensureOfficeJsCache(): Promise<{ content: string; meta: CacheMeta }> {
  if (officeJsCacheExists()) {
    const content = getCachedOfficeJs()
    const meta = getOfficeJsMeta()
    if (content && meta) return { content, meta }
  }
  const meta = await refreshOfficeJsCache()
  const content = getCachedOfficeJs() ?? STUB_OFFICE_JS
  return { content, meta }
}
