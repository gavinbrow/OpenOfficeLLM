import { describe, it, expect, afterEach } from 'vitest'
import { ExcelAdapter, columnLetter } from '../excel'
import { installExcelMock, uninstallOfficeMocks } from './mockOffice'

afterEach(() => uninstallOfficeMocks())

const HEADER = ['Region', 'Rep', 'Units', 'UnitPrice', 'Revenue']
const ROWS = [
  ['Pacific NW', 'Alvarez', 120, 45.5, 5460],
  ['Southeast', 'Okafor', 210, 38, 7980],
  ['Midwest', 'Patel', 143, 41.25, 5898.75],
]

function sheet() {
  return {
    name: 'Sales',
    values: [HEADER, ...ROWS],
    formulas: [HEADER, ...ROWS.map((r, i) => [r[0], r[1], r[2], r[3], `=C${i + 2}*D${i + 2}`])],
  }
}

describe('columnLetter', () => {
  it('maps indices to spreadsheet columns past Z', () => {
    expect(columnLetter(0)).toBe('A')
    expect(columnLetter(25)).toBe('Z')
    expect(columnLetter(26)).toBe('AA')
    expect(columnLetter(27)).toBe('AB')
  })
})

describe('ExcelAdapter.getContext', () => {
  it('serializes the used range as CSV with the header', async () => {
    installExcelMock(sheet())
    const ctx = await new ExcelAdapter().getContext('sheet')
    expect(ctx.text).toContain('Region,Rep,Units,UnitPrice,Revenue')
    expect(ctx.text).toContain('Pacific NW,Alvarez,120,45.5,5460')
  })

  it('reports the sheet list and active sheet in the outline', async () => {
    installExcelMock(sheet())
    const ctx = await new ExcelAdapter().getContext('sheet')
    expect(ctx.outline).toContain('Active sheet: Sales')
  })

  it('infers column types from the sampled rows', async () => {
    installExcelMock(sheet())
    const ctx = await new ExcelAdapter().getContext('sheet')
    const byName = Object.fromEntries((ctx.schema ?? []).map((c) => [c.name, c.type]))
    expect(byName.Region).toBe('string')
    expect(byName.Units).toBe('number')
    expect(byName.UnitPrice).toBe('number')
  })

  // Values alone hide the interesting half of a spreadsheet.
  it('surfaces formulas found in the range', async () => {
    installExcelMock(sheet())
    const ctx = await new ExcelAdapter().getContext('sheet')
    expect(ctx.text).toContain('=C2*D2')
  })

  it('reports the true row count', async () => {
    installExcelMock(sheet())
    const ctx = await new ExcelAdapter().getContext('sheet')
    expect(ctx.totalRows).toBe(4) // header + 3 data rows
  })

  it('handles an empty sheet', async () => {
    installExcelMock({ name: 'Empty', values: [], formulas: [] })
    const ctx = await new ExcelAdapter().getContext('sheet')
    expect(ctx.text).toBe('')
    expect(ctx.totalRows).toBe(0)
  })

  it('returns nothing for scope none', async () => {
    installExcelMock(sheet())
    const ctx = await new ExcelAdapter().getContext('none')
    expect(ctx.text).toBe('')
  })

  // A single-cell selection is a cursor position, not a chosen range; treating
  // it as the context would answer about one cell when the user meant the table.
  it('widens a single-cell selection to the used range', async () => {
    installExcelMock({ ...sheet(), selectedAddress: 'B3', selectedSize: [1, 1] })
    const ctx = await new ExcelAdapter().getContext('range')
    expect(ctx.text).toContain('Region,Rep,Units')
  })
})

describe('ExcelAdapter.applyEdits', () => {
  it('writes literal values to the named cells', async () => {
    const state = installExcelMock(sheet())
    const result = await new ExcelAdapter().applyEdits([
      { kind: 'setCellValues', sheet: 'Sales', cells: [{ cell: 'F2', value: 'flagged' }] },
    ])
    expect(result.ok).toBe(true)
    expect(state.writes).toContainEqual({
      address: 'F2',
      kind: 'values',
      payload: [['flagged']],
    })
  })

  it('rejects a malformed cell address before writing anything', async () => {
    const state = installExcelMock(sheet())
    const result = await new ExcelAdapter().applyEdits([
      { kind: 'setCellValues', sheet: 'Sales', cells: [{ cell: 'not a cell', value: 1 }] },
    ])
    expect(result.ok).toBe(false)
    expect(state.writes).toHaveLength(0)
  })

  it('writes a formula into an empty cell', async () => {
    const state = installExcelMock(sheet())
    const result = await new ExcelAdapter().applyEdits([
      { kind: 'setCellFormulas', sheet: 'Sales', cells: [{ cell: 'F2', formula: '=SUM(E2:E4)' }] },
    ])
    expect(result.ok).toBe(true)
    expect(state.writes).toContainEqual({
      address: 'F2',
      kind: 'formulas',
      payload: [['=SUM(E2:E4)']],
    })
  })

  // Silently overwriting someone's formula is the one destructive thing a
  // spreadsheet assistant can do that a user may not notice for months.
  it('refuses to overwrite an existing formula and says which cell', async () => {
    const state = installExcelMock(sheet())
    const result = await new ExcelAdapter().applyEdits([
      { kind: 'setCellFormulas', sheet: 'Sales', cells: [{ cell: 'E2', formula: '=999' }] },
    ])
    expect(result.ok).toBe(false)
    expect(result.summary).toContain('E2')
    expect(result.summary).toContain('=C2*D2')
    expect(state.writes).toHaveLength(0)
  })

  it('rejects a Word-only edit kind', async () => {
    installExcelMock(sheet())
    const result = await new ExcelAdapter().applyEdits([{ kind: 'replaceSelection', text: 'nope' }])
    expect(result.ok).toBe(false)
    expect(result.summary).toContain('not supported in Excel')
  })
})

describe('ExcelAdapter.search', () => {
  it('returns cell addresses for matches', async () => {
    installExcelMock(sheet())
    const hits = await new ExcelAdapter().search('Okafor')
    expect(hits).toHaveLength(1)
    expect(hits[0].location).toBe('B3')
  })

  it('returns nothing for a miss', async () => {
    installExcelMock(sheet())
    expect(await new ExcelAdapter().search('nobody')).toHaveLength(0)
  })
})
