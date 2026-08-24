// Phase 0 spike server — HTTPS on 127.0.0.1:7317, serving the test page.
//
// The point of the spike is that this server hosts the *page*, so the page's
// origin lives in the `loopback` address space. Its fetch to Ollama on :11434
// is then loopback -> loopback, which Chromium's Local Network Access check
// does not gate. Nothing here proxies Ollama — that would defeat the test.

import { createServer } from 'node:https'
import { readFileSync, existsSync } from 'node:fs'
import { extname, join, normalize, dirname } from 'node:path'
import { fileURLToPath } from 'node:url'

const here = dirname(fileURLToPath(import.meta.url))
const publicDir = join(here, 'public')
const certDir = join(here, 'certs')
const PORT = 7317
const HOST = '127.0.0.1'

if (!existsSync(join(certDir, 'server.crt'))) {
  console.error('No certs found. Run:  npm run certs')
  process.exit(1)
}

const MIME = {
  '.html': 'text/html; charset=utf-8',
  '.js': 'text/javascript; charset=utf-8',
  '.css': 'text/css; charset=utf-8',
  '.png': 'image/png',
  '.json': 'application/json; charset=utf-8',
}

const server = createServer(
  {
    key: readFileSync(join(certDir, 'server.key')),
    cert: readFileSync(join(certDir, 'server.crt')),
  },
  (req, res) => {
    const url = new URL(req.url, `https://${HOST}:${PORT}`)
    console.log(
      `${new Date().toISOString()}  ${req.method} ${url.pathname}  ua=${(req.headers['user-agent'] || '').slice(0, 60)}`,
    )

    // Server-side control probe: proves Ollama is reachable from *outside* the
    // browser, so a browser-side failure can be pinned on LNA rather than on
    // Ollama being down.
    if (url.pathname === '/control/ollama') {
      fetch('http://127.0.0.1:11434/api/tags')
        .then(async (r) => {
          res.writeHead(200, { 'content-type': 'application/json' })
          res.end(JSON.stringify({ ok: r.ok, status: r.status, body: await r.text() }))
        })
        .catch((e) => {
          res.writeHead(200, { 'content-type': 'application/json' })
          res.end(JSON.stringify({ ok: false, error: String(e) }))
        })
      return
    }

    // Collects the pane's findings so results survive even if devtools can't
    // be attached to the WebView2 frame.
    if (url.pathname === '/report' && req.method === 'POST') {
      let body = ''
      req.on('data', (c) => (body += c))
      req.on('end', () => {
        console.log('\n===== REPORT FROM PANE =====')
        console.log(body)
        console.log('===== END REPORT =====\n')
        res.writeHead(204).end()
      })
      return
    }

    const rel = url.pathname === '/' ? '/taskpane.html' : url.pathname
    const filePath = join(publicDir, normalize(rel).replace(/^(\.\.[/\\])+/, ''))
    if (!filePath.startsWith(publicDir) || !existsSync(filePath)) {
      res.writeHead(404, { 'content-type': 'text/plain' }).end('not found')
      return
    }
    res.writeHead(200, {
      'content-type': MIME[extname(filePath)] ?? 'application/octet-stream',
      'cache-control': 'no-store',
    })
    res.end(readFileSync(filePath))
  },
)

server.listen(PORT, HOST, () => {
  console.log(`spike server listening on https://${HOST}:${PORT}`)
  console.log(`task pane:  https://${HOST}:${PORT}/taskpane.html`)
})
