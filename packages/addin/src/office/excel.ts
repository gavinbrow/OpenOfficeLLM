// Excel host adapter (P4.11–P4.16).
//
// The sampling logic is the load-bearing part. A used range of 100k rows
// serialized in full is both unusable as context and a guaranteed context-length
// error, so anything past a threshold is sent as header + head + tail + inferred
// column types + the true row count. The row count matters more than it looks:
// without it a model happily reports the sample's sum as the sheet's total.

import type { ColumnSchema, ContextScope, DocumentContext, Edit } from '@openofficellm/shared'
import {
  estimateTokens,
  type ApplyResult,
  type HostAdapter,
  type SearchHit,
  type SnapshotPayload,
} from '@openofficellm/ui'

/** Rows sent verbatim before sampling kicks in. */
const FULL_ROW_LIMIT = 200
/** Rows from each end once sampling is active. */
const SAMPLE_HEAD = 60
const SAMPLE_TAIL = 20
/** Hard ceiling on cells read in one go, whatever the row count. */
const MAX_CELLS = 40_000

function cellToString(v: unknown): string {
  if (v === null || v === undefined) return ''
  if (typeof v === 'string') return v
  if (typeof v === 'number' || typeof v === 'boolean') return String(v)
  return String(v)
}

/** CSV row with the minimum quoting that survives a round trip. */
function toCsvRow(values: unknown[]): string {
  return values
    .map((v) => {
      const s = cellToString(v)
      return /[",\n]/.test(s) ? `"${s.replace(/"/g, '""')}"` : s
    })
    .join(',')
}

function inferType(values: unknown[]): ColumnSchema['type'] {
  const present = values.filter((v) => v !== null && v !== undefined && v !== '')
  if (present.length === 0) return 'null'
  let numbers = 0
  let booleans = 0
  let dates = 0
  for (const v of present) {
    if (typeof v === 'number') {
      numbers++
      continue
    }
    if (typeof v === 'boolean') {
      booleans++
      continue
    }
    const s = String(v)
    if (/^\d{4}-\d{2}-\d{2}/.test(s) || /^\d{1,2}\/\d{1,2}\/\d{2,4}$/.test(s)) dates++
  }
  if (numbers === present.length) return 'number'
  if (booleans === present.length) return 'boolean'
  if (dates === present.length) return 'date'
  if (numbers > 0 && numbers < present.length) return 'string' // mixed
  return 'string'
}

function buildSchema(header: unknown[], rows: unknown[][]): ColumnSchema[] {
  return header.map((name, col) => {
    const column = rows.map((r) => r[col])
    return {
      name: cellToString(name) || `Column ${col + 1}`,
      type: inferType(column),
      sampleValues: column
        .filter((v) => v !== null && v !== undefined && v !== '')
        .slice(0, 3)
        .map((v) => cellToString(v)),
    }
  })
}

export class ExcelAdapter implements HostAdapter {
  readonly host = 'excel' as const

  async getContext(scope: ContextScope): Promise<DocumentContext> {
    if (scope === 'none') {
      return { host: 'excel', scope, text: '', tokenEstimate: 0 }
    }

    return Excel.run(async (context) => {
      const workbook = context.workbook
      const sheets = workbook.worksheets
      sheets.load('items/name')
      const sheet = sheets.getActiveWorksheet()
      sheet.load('name')

      const selected = workbook.getSelectedRange()
      selected.load('address,rowCount,columnCount')
      await context.sync()

      const outline = [
        `Active sheet: ${sheet.name}`,
        `Sheets: ${sheets.items.map((s) => s.name).join(', ')}`,
        `Selection: ${selected.address}`,
      ].join('\n')

      // A one-cell selection is a cursor position, not a chosen range. Treating
      // it as the context would answer questions about a single cell when the
      // user meant the table they are sitting in.
      const meaningfulSelection = selected.rowCount * selected.columnCount > 1
      const target = scope === 'range' && meaningfulSelection ? selected : sheet.getUsedRange(true)

      target.load('address,rowCount,columnCount')
      await context.sync()

      const totalRows = target.rowCount ?? 0
      const totalCols = target.columnCount ?? 0
      if (totalRows === 0 || totalCols === 0) {
        return {
          host: 'excel' as const,
          scope,
          text: '',
          outline,
          totalRows: 0,
          tokenEstimate: estimateTokens(outline),
        }
      }

      const sampled = totalRows > FULL_ROW_LIMIT || totalRows * totalCols > MAX_CELLS
      let headerRow: unknown[] = []
      let bodyRows: unknown[][] = []
      let formulaRows: unknown[][] = []

      if (!sampled) {
        target.load('values,formulas')
        await context.sync()
        const values = (target.values ?? []) as unknown[][]
        headerRow = values[0] ?? []
        bodyRows = values.slice(1)
        formulaRows = (target.formulas ?? []) as unknown[][]
      } else {
        // Read only the slices we will actually send. Loading the full range and
        // slicing in JS defeats the point — the cost is in the marshalling.
        const head = target.getRow(0).getResizedRange(Math.min(SAMPLE_HEAD, totalRows - 1), 0)
        const tail = target
          .getRow(Math.max(0, totalRows - SAMPLE_TAIL))
          .getResizedRange(Math.min(SAMPLE_TAIL - 1, totalRows - 1), 0)
        head.load('values,formulas')
        tail.load('values')
        await context.sync()
        const headValues = (head.values ?? []) as unknown[][]
        headerRow = headValues[0] ?? []
        bodyRows = headValues.slice(1)
        formulaRows = (head.formulas ?? []) as unknown[][]
        const tailValues = (tail.values ?? []) as unknown[][]
        if (tailValues.length > 0) {
          bodyRows = [...bodyRows, ['…'], ...tailValues]
        }
      }

      const csv = [toCsvRow(headerRow), ...bodyRows.map((r) => toCsvRow(r))].join('\n')

      // Formulas are the interesting part of a spreadsheet and values alone
      // hide them. Include the distinct ones seen in the sample.
      const formulas = new Set<string>()
      formulaRows.forEach((row, rIdx) => {
        row.forEach((f, cIdx) => {
          const s = cellToString(f)
          if (s.startsWith('=')) {
            formulas.add(`${columnLetter(cIdx)}${rIdx + 1}: ${s}`)
          }
        })
      })
      const formulaNote =
        formulas.size > 0
          ? `\n\nFormulas in the sampled rows:\n${[...formulas].slice(0, 40).join('\n')}`
          : ''

      const rangeNote = sampled
        ? `\n\n[showing ${SAMPLE_HEAD} rows from the top and ${SAMPLE_TAIL} from the bottom of ${totalRows} total]`
        : ''

      const text = `Range ${target.address}\n${csv}${formulaNote}${rangeNote}`
      return {
        host: 'excel' as const,
        scope,
        text,
        outline,
        schema: buildSchema(headerRow, bodyRows),
        totalRows,
        tokenEstimate: estimateTokens(text),
      }
    })
  }

  async applyEdits(edits: Edit[]): Promise<ApplyResult> {
    if (edits.length === 0) return { ok: true, summary: 'No changes to apply.' }

    return Excel.run(async (context) => {
      const applied: string[] = []
      for (const edit of edits) {
        switch (edit.kind) {
          case 'setCellValues': {
            const sheet = resolveSheet(context, edit.sheet)
            for (const cell of edit.cells) {
              if (!isValidAddress(cell.cell)) {
                return { ok: false, summary: `Invalid cell address: ${cell.cell}` }
              }
              sheet.getRange(cell.cell).values = [[cell.value]]
            }
            applied.push(`set ${edit.cells.length} value(s) on ${edit.sheet || 'the active sheet'}`)
            break
          }
          case 'setCellFormulas': {
            const sheet = resolveSheet(context, edit.sheet)
            // Refuse to clobber an existing formula unless the model is
            // overwriting the same cell it just read (P4.16). Read first.
            const targets = edit.cells.filter((c) => isValidAddress(c.cell))
            if (targets.length !== edit.cells.length) {
              return { ok: false, summary: 'One or more cell addresses were invalid.' }
            }
            const ranges = targets.map((c) => {
              const r = sheet.getRange(c.cell)
              r.load('formulas')
              return r
            })
            await context.sync()
            const clobbered = ranges
              .map((r, i) => ({
                addr: targets[i].cell,
                existing: cellToString(r.formulas?.[0]?.[0]),
              }))
              .filter((x) => x.existing.startsWith('='))
            if (clobbered.length > 0) {
              return {
                ok: false,
                summary: `Refused: ${clobbered
                  .map((c) => `${c.addr} already contains ${c.existing}`)
                  .join(', ')}. Ask the user to confirm before overwriting existing formulas.`,
              }
            }
            targets.forEach((c, i) => {
              ranges[i].formulas = [[c.formula]]
            })
            applied.push(`wrote ${targets.length} formula(s)`)
            break
          }
          case 'applyFormatting': {
            const sheet = context.workbook.worksheets.getActiveWorksheet()
            if (!isValidAddress(edit.rangeId)) {
              return { ok: false, summary: `Invalid range: ${edit.rangeId}` }
            }
            const range = sheet.getRange(edit.rangeId)
            const f = edit.formatting as Record<string, unknown>
            if (typeof f.bold === 'boolean') range.format.font.bold = f.bold
            if (typeof f.italic === 'boolean') range.format.font.italic = f.italic
            if (typeof f.color === 'string') range.format.font.color = f.color
            if (typeof f.fill === 'string') range.format.fill.color = f.fill
            if (typeof f.numberFormat === 'string') {
              range.numberFormat = [[f.numberFormat]]
            }
            applied.push(`formatted ${edit.rangeId}`)
            break
          }
          default:
            return { ok: false, summary: `${edit.kind} is not supported in Excel.` }
        }
      }
      await context.sync()
      return { ok: true, summary: `Applied: ${applied.join('; ')}.` }
    })
  }

  async snapshot(): Promise<SnapshotPayload> {
    return Excel.run(async (context) => {
      const sheet = context.workbook.worksheets.getActiveWorksheet()
      sheet.load('name')
      const used = sheet.getUsedRange(true)
      used.load('address,values,formulas,rowCount,columnCount')
      await context.sync()

      const cells = (used.rowCount ?? 0) * (used.columnCount ?? 0)
      return {
        id: `snap_${Date.now().toString(36)}`,
        host: 'excel' as const,
        createdAt: new Date().toISOString(),
        sizeBytes: cells * 16,
        data: {
          kind: 'range' as const,
          sheet: sheet.name,
          address: used.address,
          values: used.values,
          formulas: used.formulas,
        },
      }
    })
  }

  async restore(snapshot: SnapshotPayload): Promise<void> {
    const data = snapshot.data as {
      kind: string
      sheet: string
      address: string
      formulas: unknown[][]
    }
    if (data?.kind !== 'range') throw new Error('Snapshot cannot be restored.')
    await Excel.run(async (context) => {
      const sheet = context.workbook.worksheets.getItem(data.sheet)
      // Restore formulas rather than values: writing back computed values would
      // turn every formula in the range into a literal.
      sheet.getRange(data.address).formulas = data.formulas as string[][]
      await context.sync()
    })
  }

  async search(query: string, limit = 20): Promise<SearchHit[]> {
    if (!query.trim()) return []
    return Excel.run(async (context) => {
      const sheet = context.workbook.worksheets.getActiveWorksheet()
      const used = sheet.getUsedRange(true)
      used.load('values,address,rowIndex,columnIndex')
      await context.sync()
      const values = (used.values ?? []) as unknown[][]
      const needle = query.toLowerCase()
      const hits: SearchHit[] = []
      const baseRow = used.rowIndex ?? 0
      const baseCol = used.columnIndex ?? 0
      for (let r = 0; r < values.length && hits.length < limit; r++) {
        for (let c = 0; c < values[r].length && hits.length < limit; c++) {
          const s = cellToString(values[r][c])
          if (s && s.toLowerCase().includes(needle)) {
            hits.push({
              location: `${columnLetter(baseCol + c)}${baseRow + r + 1}`,
              text: s.slice(0, 300),
            })
          }
        }
      }
      return hits
    })
  }
}

function resolveSheet(context: Excel.RequestContext, name?: string): Excel.Worksheet {
  if (name && name.trim()) return context.workbook.worksheets.getItem(name)
  return context.workbook.worksheets.getActiveWorksheet()
}

/** A1, $B$2, A1:C10, Sheet-less only — the sheet comes from the edit. */
function isValidAddress(addr: string): boolean {
  return /^\$?[A-Za-z]{1,3}\$?\d{1,7}(:\$?[A-Za-z]{1,3}\$?\d{1,7})?$/.test(addr.trim())
}

export function columnLetter(index: number): string {
  let n = index
  let out = ''
  do {
    out = String.fromCharCode(65 + (n % 26)) + out
    n = Math.floor(n / 26) - 1
  } while (n >= 0)
  return out
}
