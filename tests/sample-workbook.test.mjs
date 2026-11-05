import assert from 'node:assert/strict'
import { readFileSync } from 'node:fs'
import { test } from 'node:test'
import vm from 'node:vm'

// Reads sample/sample-equipment.xlsx with the vendored SheetJS build and runs it
// through the same parse and model steps the app uses after a file is dropped.
function loadContext() {
  const html = readFileSync(new URL('../index.html', import.meta.url), 'utf8')
  const pipeline = html.match(/<script id="pipeline">([\s\S]*?)<\/script>/)
  assert.ok(pipeline, 'pipeline script is present')
  const context = { console }
  vm.createContext(context)
  vm.runInContext(readFileSync(new URL('../vendor/xlsx.bundle.js', import.meta.url), 'utf8'), context)
  vm.runInContext(pipeline[1], context)
  return context
}

function matchSheet(names, needle) {
  const flat = (s) => String(s).toLowerCase().replace(/[^a-z0-9]/g, '')
  return names.find((name) => flat(name) === flat(needle))
}

test('sample workbook loads through the app pipeline', () => {
  const ctx = loadContext()
  const bytes = readFileSync(new URL('../sample/sample-equipment.xlsx', import.meta.url))
  const wb = ctx.XLSX.read(new Uint8Array(bytes), { type: 'array', dense: true, cellHTML: false })
  const aoa = (needle) => {
    const name = matchSheet(wb.SheetNames, needle)
    assert.ok(name, `${needle} tab is present`)
    return ctx.XLSX.utils.sheet_to_json(wb.Sheets[name], { header: 1, raw: false, defval: '' })
  }

  let esr = ctx.parseEquipmentStatus(aoa('Equipment Status Report'), { catAssign: {}, ignoreStepValues: [], colOverride: {} })
  const catAssign = {}
  esr.stepCols.forEach((col) => { catAssign[col.header.toLowerCase()] = col.defaultCats.slice() })
  esr = ctx.parseEquipmentStatus(aoa('Equipment Status Report'), { catAssign, ignoreStepValues: [], colOverride: {} })

  const model = ctx.buildModel(
    esr,
    ctx.parseDropdowns(aoa('Drop Downs- VF')),
    ctx.parseIssues(aoa('Issue Grid Report')),
    null,
    null,
  )

  assert.equal(model.totals.equipment, 1200)
  assert.equal(model.totals.systems, 40)
  assert.equal(model.totals.issues, 3000)
  assert.equal(model.totals.unmatchedIssues, 1)
  assert.ok(esr.stepCols.length > 0, 'step columns are detected')
})
