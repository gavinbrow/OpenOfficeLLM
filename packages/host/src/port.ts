import net from 'node:net'
import { DEFAULT_PORT } from './paths.js'

export interface PortSelectionResult {
  port: number
  scanned: boolean
  tried: number[]
}

function tryPort(port: number, host = '127.0.0.1'): Promise<boolean> {
  return new Promise((resolve) => {
    const srv = net.createServer()
    srv.unref()
    srv.once('error', () => resolve(false))
    srv.listen(port, host, () => {
      srv.close(() => resolve(true))
    })
  })
}

export async function selectPort(
  preferred: number | undefined,
  maxScan = 50,
): Promise<PortSelectionResult> {
  const base = preferred ?? DEFAULT_PORT
  const tried: number[] = []
  for (let i = 0; i <= maxScan; i++) {
    const p = base + i
    if (p > 65535) break
    tried.push(p)
    const free = await tryPort(p)
    if (free) {
      return { port: p, scanned: i > 0, tried }
    }
  }
  throw new Error(
    `No free port found in range ${base}..${Math.min(base + maxScan, 65535)} ` +
      `(tried ${tried.join(', ')}). Use --port <n> to pick another.`,
  )
}
