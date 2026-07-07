import assert from 'node:assert/strict'
import { readFileSync } from 'node:fs'
import { test } from 'node:test'
import vm from 'node:vm'

function loadPipeline() {
  const html = readFileSync(new URL('../index.html', import.meta.url), 'utf8')
  const match = html.match(/<script id="pipeline">([\s\S]*?)<\/script>/)
  assert.ok(match, 'pipeline script is present')
  const context = { console }
  vm.createContext(context)
  vm.runInContext(match[1], context)
  return context
}

const equipment = [
  {
    systemName: 'Alpha',
    discipline: 'Mechanical',
    building: 'B1',
    milestone: 'M1',
    classification: 'Pump',
    present: { QAQC: false, DV: false, EHS: false },
    categories: [],
    score: 0,
    issueCount: 0,
    hasIssues: false,
  },
  {
    systemName: 'Alpha',
    discipline: 'Mechanical',
    building: 'B1',
    milestone: 'M1',
    classification: 'Pump',
    present: { QAQC: true, DV: false, EHS: false },
    categories: ['QAQC'],
    score: 1,
    issueCount: 1,
    hasIssues: true,
  },
  {
    systemName: 'Beta',
    discipline: 'Electrical',
    building: 'B2',
    milestone: 'M2',
    classification: 'Panel',
    present: { QAQC: true, DV: true, EHS: false },
    categories: ['QAQC', 'DV'],
    score: 2,
    issueCount: 2,
    hasIssues: true,
  },
  {
    systemName: 'Beta',
    discipline: 'Electrical',
    building: 'B2',
    milestone: 'M2',
    classification: 'Panel',
    present: { QAQC: true, DV: true, EHS: true },
    categories: ['QAQC', 'DV', 'EHS'],
    score: 3,
    issueCount: 4,
    hasIssues: true,
  },
]

test('step selections distinguish no filter, none, any, and all', () => {
  const ctx = loadPipeline()
  const noFilter = { none: false, categories: [], mode: 'any' }
  const none = { none: true, categories: [], mode: 'any' }
  const any = { none: false, categories: ['QAQC', 'EHS'], mode: 'any' }
  const all = { none: false, categories: ['QAQC', 'EHS'], mode: 'all' }

  assert.deepEqual(equipment.filter((e) => ctx.matchStepSelection(e, noFilter)).map((e) => e.score), [0, 1, 2, 3])
  assert.deepEqual(equipment.filter((e) => ctx.matchStepSelection(e, none)).map((e) => e.score), [0])
  assert.deepEqual(equipment.filter((e) => ctx.matchStepSelection(e, any)).map((e) => e.score), [1, 2, 3])
  assert.deepEqual(equipment.filter((e) => ctx.matchStepSelection(e, all)).map((e) => e.score), [3])
})

test('coverage breakdown returns score ratios for each dimension group', () => {
  const ctx = loadPipeline()
  const rows = ctx.coverageBreakdownBy(equipment, (e) => e.systemName)
  const alpha = rows.find((row) => row.key === 'Alpha')
  const beta = rows.find((row) => row.key === 'Beta')

  assert.equal(alpha.n, 2)
  assert.equal(alpha.scoreDist[0], 1)
  assert.equal(alpha.scoreDist[1], 1)
  assert.equal(alpha.scorePct[0], 50)
  assert.equal(alpha.scorePct[1], 50)
  assert.equal(beta.scoreDist[2], 1)
  assert.equal(beta.scoreDist[3], 1)
  assert.equal(beta.scorePct[2], 50)
  assert.equal(beta.scorePct[3], 50)
})

test('coverage tier breakdowns can group by steps or a selected dimension', () => {
  const ctx = loadPipeline()
  const bySteps = ctx.tierBreakdown(equipment, 2, 'Steps')
  const bySystem = ctx.tierBreakdown(equipment, 3, 'System')

  assert.equal(bySteps[0].key, 'QAQC + DV')
  assert.equal(bySteps[0].ratio, 100)
  assert.equal(bySystem[0].key, 'Beta')
  assert.equal(bySystem[0].ratio, 100)
  assert.equal(bySystem[0].score, 3)
})

test('blind-spot risk scoring produces finite bars from coverage rows', () => {
  const ctx = loadPipeline()
  const areas = ctx.scoreRisk(ctx.areaStatsFrom(equipment, (e) => e.systemName), 'blind')
  const alpha = areas.find((row) => row.key === 'Alpha')

  assert.ok(Number.isFinite(alpha.risk))
  assert.ok(alpha.risk > 0)
  assert.equal(alpha.issues, 1)
  assert.equal(alpha.scoreDist[0], 1)
  assert.equal(alpha.scoreDist[1], 1)
})
