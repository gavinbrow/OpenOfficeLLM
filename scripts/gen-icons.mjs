// Generates the app icons every shell's manifest points at.
//
// Both manifests need real PNGs and neither tolerates a miss: Office fetches
// them while building the ribbon and can drop an add-in whose icons 404, and
// Chrome refuses to load an extension whose declared icons are absent.
//
// Writing the PNGs by hand keeps a raster-image dependency out of the build for
// what is four flat-colour glyphs. Run `npm run gen:icons` in a package after
// changing BRAND, then commit the output — no build step regenerates them.
//
// Usage: node scripts/gen-icons.mjs --out <dir> [--sizes 16,32,128]

import { deflateSync } from 'node:zlib'
import { writeFileSync, mkdirSync } from 'node:fs'
import { join, resolve } from 'node:path'
import { parseArgs } from 'node:util'

const { values } = parseArgs({
  options: {
    out: { type: 'string' },
    sizes: { type: 'string', default: '16,32,64,80' },
  },
})

if (!values.out) {
  console.error('gen-icons: --out <dir> is required')
  process.exit(2)
}

const outDir = resolve(process.cwd(), values.out)
mkdirSync(outDir, { recursive: true })

const SIZES = values.sizes
  .split(',')
  .map((s) => Number(s.trim()))
  .filter((n) => Number.isInteger(n) && n > 0)

if (SIZES.length === 0) {
  console.error(`gen-icons: no usable sizes in "${values.sizes}"`)
  process.exit(2)
}

/** Tailwind `accent.DEFAULT` — keep in sync with packages/ui/tailwind-preset.js. */
const BRAND = [0x0f, 0x6c, 0xbd]
/** Supersampling factor. 16px is small enough that hard edges look broken. */
const SS = 4

const crcTable = Array.from({ length: 256 }, (_, n) => {
  let c = n
  for (let k = 0; k < 8; k++) c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1
  return c >>> 0
})

function crc32(buf) {
  let c = 0xffffffff
  for (const b of buf) c = crcTable[(c ^ b) & 0xff] ^ (c >>> 8)
  return (c ^ 0xffffffff) >>> 0
}

function chunk(type, data) {
  const len = Buffer.alloc(4)
  len.writeUInt32BE(data.length)
  const body = Buffer.concat([Buffer.from(type, 'ascii'), data])
  const crc = Buffer.alloc(4)
  crc.writeUInt32BE(crc32(body))
  return Buffer.concat([len, body, crc])
}

/** Rounded square covering most of the tile, in normalized 0..1 space. */
function inRoundedSquare(x, y) {
  const inset = 0.06
  const r = 0.22
  const lo = inset
  const hi = 1 - inset
  if (x < lo || x > hi || y < lo || y > hi) return false
  const dx = Math.max(lo + r - x, 0, x - (hi - r))
  const dy = Math.max(lo + r - y, 0, y - (hi - r))
  return dx * dx + dy * dy <= r * r
}

/** Four-point sparkle (an astroid), the conventional "AI" glyph. Concave
 *  sides read as a star rather than a diamond even at 16px. */
function inSparkle(x, y) {
  const cx = 0.5
  const cy = 0.5
  const rad = 0.34
  const dx = Math.abs(x - cx) / rad
  const dy = Math.abs(y - cy) / rad
  if (dx > 1 || dy > 1) return false
  return Math.sqrt(dx) + Math.sqrt(dy) <= 1
}

function png(size) {
  const ihdr = Buffer.alloc(13)
  ihdr.writeUInt32BE(size, 0)
  ihdr.writeUInt32BE(size, 4)
  ihdr[8] = 8 // bit depth
  ihdr[9] = 6 // colour type: RGBA
  // 10..12 = compression / filter / interlace, all 0

  const raw = Buffer.alloc(size * (size * 4 + 1))
  let p = 0
  const samples = SS * SS
  for (let y = 0; y < size; y++) {
    raw[p++] = 0 // filter: none
    for (let x = 0; x < size; x++) {
      // Coverage of the tile and of the glyph, sampled independently so the
      // white sparkle antialiases against the brand fill rather than against
      // transparency.
      let bg = 0
      let fg = 0
      for (let sy = 0; sy < SS; sy++) {
        for (let sx = 0; sx < SS; sx++) {
          const u = (x + (sx + 0.5) / SS) / size
          const v = (y + (sy + 0.5) / SS) / size
          if (inRoundedSquare(u, v)) bg++
          if (inSparkle(u, v)) fg++
        }
      }
      if (bg === 0) {
        raw[p++] = 0
        raw[p++] = 0
        raw[p++] = 0
        raw[p++] = 0
        continue
      }
      const alpha = bg / samples
      // Sparkle only shows where the tile is; premultiplied blend toward white.
      const mix = Math.min(fg / samples, alpha) / alpha
      raw[p++] = Math.round(BRAND[0] + (255 - BRAND[0]) * mix)
      raw[p++] = Math.round(BRAND[1] + (255 - BRAND[1]) * mix)
      raw[p++] = Math.round(BRAND[2] + (255 - BRAND[2]) * mix)
      raw[p++] = Math.round(alpha * 255)
    }
  }

  return Buffer.concat([
    Buffer.from([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]),
    chunk('IHDR', ihdr),
    chunk('IDAT', deflateSync(raw, { level: 9 })),
    chunk('IEND', Buffer.alloc(0)),
  ])
}

for (const size of SIZES) {
  const file = join(outDir, `icon-${size}.png`)
  writeFileSync(file, png(size))
  console.log(`wrote ${file}`)
}
