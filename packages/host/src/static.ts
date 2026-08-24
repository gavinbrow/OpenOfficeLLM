import fs from 'node:fs'
import path from 'node:path'
import type { Context } from 'hono'
import { logger } from './logging.js'

const MIME: Record<string, string> = {
  '.html': 'text/html; charset=utf-8',
  '.htm': 'text/html; charset=utf-8',
  '.js': 'text/javascript; charset=utf-8',
  '.mjs': 'text/javascript; charset=utf-8',
  '.css': 'text/css; charset=utf-8',
  '.json': 'application/json; charset=utf-8',
  '.map': 'application/json; charset=utf-8',
  '.png': 'image/png',
  '.jpg': 'image/jpeg',
  '.jpeg': 'image/jpeg',
  '.gif': 'image/gif',
  '.svg': 'image/svg+xml',
  '.ico': 'image/x-icon',
  '.webp': 'image/webp',
  '.woff': 'font/woff',
  '.woff2': 'font/woff2',
  '.ttf': 'font/ttf',
  '.otf': 'font/otf',
  '.txt': 'text/plain; charset=utf-8',
  '.pdf': 'application/pdf',
  '.wasm': 'application/wasm',
  '.xml': 'application/xml; charset=utf-8',
}

const META_TOKEN_REPLACEMENT = '<meta name="auth-token" content="">'
// Match the CDN script tag by regex so we tolerate attribute additions (defer,
// integrity, etc.) and the .debug variant. The host rewrites the tag to the
// local cache when the CDN was unreachable at startup.
const OFFICE_JS_CDN_RE =
  /<script\s+src=["']https:\/\/appsforoffice\.microsoft\.com\/lib\/1\/hosted\/office(?:\.debug)?\.js["']/i
const OFFICE_JS_LOCAL_TAG = '<script src="/office/office.js"'

/** Filenames that count as the app shell and therefore need token injection
 *  and no-cache headers. Vite emits `index.html`; the spike used
 *  `taskpane.html`. Accept both.
 *
 *  commands.html is deliberately NOT here. It is the manifest's FunctionFile
 *  and carries no app code, so it has no use for the token — and every page
 *  the token is stamped into is another place it can leak from. It still gets
 *  the Office.js rewrite below, which applies to all HTML. */
const APP_shell_BASENAMES = new Set(['index.html', 'taskpane.html'])

export interface StaticServeOptions {
  rootDir: string
  token: string
  officeJsLocalAvailable: boolean
}

export function resolveStaticPath(rootDir: string, requestPath: string): string | null {
  const rel = requestPath === '/' ? '/index.html' : requestPath
  const clean = rel.split('?')[0].split('#')[0]
  // decodeURIComponent throws URIError on malformed escapes ("/%ZZ"). Left
  // uncaught it escapes to the global handler and answers 500, which both leaks
  // that the path reached the file layer and turns a bad request into a
  // server error. Treat undecodable paths as simply not found.
  let decoded: string
  try {
    decoded = decodeURIComponent(clean)
  } catch {
    return null
  }
  const normalized = path.normalize(decoded).replace(/^([/\\])+/, '')
  const full = path.join(rootDir, normalized)
  const rootResolved = path.resolve(rootDir)
  const fullResolved = path.resolve(full)
  const relFromRoot = path.relative(rootResolved, fullResolved)
  if (relFromRoot.startsWith('..') || path.isAbsolute(relFromRoot)) {
    return null
  }
  return fullResolved
}

function isAppShell(filePath: string): boolean {
  return APP_shell_BASENAMES.has(path.basename(filePath).toLowerCase())
}

function injectToken(html: string, token: string): string {
  const replacement = `<meta name="auth-token" content="${escapeHtmlAttr(token)}">`
  if (html.includes(META_TOKEN_REPLACEMENT)) {
    return html.replace(META_TOKEN_REPLACEMENT, replacement)
  }
  if (/<meta\s+name=["']auth-token["']/i.test(html)) {
    return html.replace(
      /<meta\s+name=["']auth-token["']\s+content=["'][^"']*["']\s*\/?>/i,
      replacement,
    )
  }
  return html.replace(/<head>(?!\s*<meta\s+name=["']auth-token["'])/i, `<head>${replacement}`)
}

function escapeHtmlAttr(value: string): string {
  return value
    .replace(/&/g, '&amp;')
    .replace(/"/g, '&quot;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
}

function rewriteOfficeJs(html: string, useLocal: boolean): string {
  if (!useLocal) return html
  return html.replace(OFFICE_JS_CDN_RE, OFFICE_JS_LOCAL_TAG)
}

export function serveStatic(c: Context, opts: StaticServeOptions): Response | null {
  const reqPath = c.req.path
  const full = resolveStaticPath(opts.rootDir, reqPath)
  if (!full) {
    return new Response('not found', { status: 404, headers: { 'content-type': 'text/plain' } })
  }
  let stat: fs.Stats
  try {
    stat = fs.statSync(full)
  } catch {
    return new Response('not found', { status: 404, headers: { 'content-type': 'text/plain' } })
  }
  if (stat.isDirectory()) {
    const index = path.join(full, 'index.html')
    try {
      fs.statSync(index)
      return serveFile(index, opts, true)
    } catch {
      return new Response('forbidden', { status: 403, headers: { 'content-type': 'text/plain' } })
    }
  }
  return serveFile(full, opts, isAppShell(full))
}

function serveFile(full: string, opts: StaticServeOptions, isTaskpane: boolean): Response {
  const ext = path.extname(full).toLowerCase()
  const mime = MIME[ext] ?? 'application/octet-stream'
  const data = fs.readFileSync(full)
  const isHtml = ext === '.html' || ext === '.htm'
  if (isHtml) {
    let html = data.toString('utf8')
    if (isTaskpane) html = injectToken(html, opts.token)
    html = rewriteOfficeJs(html, opts.officeJsLocalAvailable)
    return new Response(html, {
      status: 200,
      headers: {
        'content-type': 'text/html; charset=utf-8',
        // The app shell must never be cached: the token is stamped into it at
        // request time and rotates every launch. Other HTML (commands.html)
        // rides along with the same headers — it is one small file loaded once
        // per ribbon build, so there is nothing to gain from caching it.
        'cache-control': 'no-store, no-cache, must-revalidate',
        pragma: 'no-cache',
      },
    })
  }
  return new Response(data, {
    status: 200,
    headers: {
      'content-type': mime,
      'cache-control': 'public, max-age=3600',
    },
  })
}

export function staticDirExists(rootDir: string): boolean {
  try {
    const s = fs.statSync(rootDir)
    return s.isDirectory()
  } catch {
    return false
  }
}

export function warnIfMissing(rootDir: string): void {
  if (!staticDirExists(rootDir)) {
    logger.warn({
      msg: 'addin dist directory missing; /api/* still serves, task pane will not load',
      dir: rootDir,
    })
  }
}
