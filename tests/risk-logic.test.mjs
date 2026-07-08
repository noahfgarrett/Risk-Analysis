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

test('issue distribution rows include share labels for chart bars', () => {
  const ctx = loadPipeline()
  const rows = ctx.issueShareRows([
    { key: 'Alpha', issueCount: 1 },
    { key: 'Beta', issueCount: 3 },
  ])

  assert.equal(rows[0].issueShare, 25)
  assert.equal(rows[0].issueShareLabel, '25%')
  assert.equal(rows[1].issueShare, 75)
  assert.equal(rows[1].issueShareLabel, '75%')
})

test('step issue contribution ranks QAQC DV and EHS by caught issues', () => {
  const ctx = loadPipeline()
  const rows = ctx.stepIssueContributions(equipment)

  assert.equal(JSON.stringify(rows.map((row) => row.step)), JSON.stringify(['QAQC', 'DV', 'EHS']))
  assert.equal(JSON.stringify(rows.map((row) => row.issueCount)), JSON.stringify([7, 6, 4]))
  assert.equal(JSON.stringify(rows.map((row) => row.ratio)), JSON.stringify([41, 35, 24]))
})

test('update asset selection prefers the versioned plain HTML asset for simple downloads', () => {
  const ctx = loadPipeline()
  const selected = ctx.selectUpdateAsset([
    { name: 'Risk-Analysis.html.gz', url: 'api-gzip', browser_download_url: 'download-gzip' },
    { name: 'Risk-Analysis.html', url: 'api-html', browser_download_url: 'download-html' },
    { name: 'Risk-Analysis-v1.0.4.html', url: 'api-versioned', browser_download_url: 'download-versioned' },
  ])

  assert.equal(selected.downloadKind, 'html')
  assert.equal(selected.assetName, 'Risk-Analysis-v1.0.4.html')
  assert.equal(selected.downloadUrl, 'download-versioned')
})

test('room filtering removes equipment whose system or UPN is RR', () => {
  const ctx = loadPipeline()
  const rows = [
    { name: 'keep', systemName: 'Alpha', upnRaw: '1001', upn: '1001' },
    { name: 'room system', systemName: 'RR', upnRaw: '1002', upn: '1002' },
    { name: 'room prefix', systemName: 'RR - Room Systems', upnRaw: '1003', upn: '1003' },
    { name: 'room upn', systemName: 'Beta', upnRaw: 'RR', upn: '' },
  ]

  assert.deepEqual(ctx.filterRoomEquipment(rows, true).map((row) => row.name), ['keep'])
  assert.equal(ctx.filterRoomEquipment(rows, false).length, 4)
})

test('issue distribution rows honor room and step filters', () => {
  const ctx = loadPipeline()
  const rows = ctx.issueDistributionRows([
    ...equipment,
    {
      systemName: 'RR',
      upnRaw: 'RR',
      discipline: 'Rooms',
      building: 'B3',
      milestone: 'M3',
      classification: 'Room',
      present: { QAQC: true, DV: false, EHS: false },
      categories: ['QAQC'],
      score: 1,
      issueCount: 50,
      hasIssues: true,
    },
  ], 'System', {
    excludeRooms: true,
    stepSelection: { none: false, categories: ['QAQC'], mode: 'any' },
  })

  assert.equal(JSON.stringify(rows.map((row) => row.key)), JSON.stringify(['Beta', 'Alpha']))
  assert.equal(JSON.stringify(rows.map((row) => row.issueCount)), JSON.stringify([6, 1]))
})

test('issue parsing keeps Cx Step, Category, Root Cause, and UPN per issue row', () => {
  const ctx = loadPipeline()
  const parsed = ctx.parseIssues([
    ['Issue ID', 'Equipment Name', 'UPN Tag', 'Cx Step', 'Category', 'Root Cause'],
    ['ISS-1', 'Pump A', '1001', 'Functional Test', 'COR', 'Design'],
    ['ISS-2', 'Pump A', '1001', 'Functional Test', 'NCR-D', 'Install'],
    ['ISS-3', 'Panel B', '2002', 'Startup', 'WI', 'Design'],
  ])

  assert.equal(parsed.rows.length, 3)
  assert.equal(parsed.cxStepCol, 3)
  assert.equal(parsed.categoryCol, 4)
  assert.equal(parsed.rootCauseCol, 5)
  assert.equal(JSON.stringify(parsed.rows[0]), JSON.stringify({
    equipmentName: 'Pump A',
    equipmentKey: 'pump a',
    upn: '1001',
    upnRaw: '1001',
    cxStep: 'Functional Test',
    category: 'COR',
    rootCause: 'Design',
  }))
})

test('issue distribution can count issues by root cause from matched issue rows', () => {
  const ctx = loadPipeline()
  const model = ctx.buildModel({
    equipment: [
      { name: 'Pump A', upn: '1001', upnRaw: '1001', systemName: 'Water', discipline: 'Mechanical', building: 'B1', milestone: 'M1', classification: 'Pump', score: 1, present: { QAQC: true, DV: false, EHS: false }, categories: ['QAQC'] },
      { name: 'Panel B', upn: '2002', upnRaw: '2002', systemName: 'Power', discipline: 'Electrical', building: 'B2', milestone: 'M2', classification: 'Panel', score: 2, present: { QAQC: true, DV: true, EHS: false }, categories: ['QAQC', 'DV'] },
    ],
  }, { map: { 1001: '1001 - Water System', 2002: '2002 - Power System' } }, ctx.parseIssues([
    ['Equipment Name', 'UPN Tag', 'Cx Step', 'Category', 'Root Cause'],
    ['Pump A', '1001', 'Functional Test', 'COR', 'Design'],
    ['Pump A', '1001', 'Functional Test', 'NCR-D', 'Install'],
    ['Panel B', '2002', 'Startup', 'WI', 'Design'],
  ]), { byId: new Map([['pump a', 'Pump'], ['panel b', 'Panel']]) })

  const rows = ctx.issueDistributionRowsFromIssues(model.equipment, model.issueRows, 'System', { rootCause: 'Design' })

  assert.equal(JSON.stringify(rows.map((row) => row.key)), JSON.stringify(['1001 - Water System', '2002 - Power System']))
  assert.equal(JSON.stringify(rows.map((row) => row.issueCount)), JSON.stringify([1, 1]))
  assert.equal(JSON.stringify(rows.map((row) => row.equipmentCount)), JSON.stringify([1, 1]))
})

test('issues caught by Cx Step roll up stacked category shares', () => {
  const ctx = loadPipeline()
  const model = ctx.buildModel({
    equipment: [
      { name: 'Pump A', upn: '1001', upnRaw: '1001', systemName: 'Water', discipline: 'Mechanical', building: 'B1', milestone: 'M1', classification: 'Pump', score: 1, present: { QAQC: true, DV: false, EHS: false }, categories: ['QAQC'] },
      { name: 'Panel B', upn: '2002', upnRaw: '2002', systemName: 'Power', discipline: 'Electrical', building: 'B2', milestone: 'M2', classification: 'Panel', score: 2, present: { QAQC: true, DV: true, EHS: false }, categories: ['QAQC', 'DV'] },
    ],
  }, { map: { 1001: '1001 - Water System', 2002: '2002 - Power System' } }, ctx.parseIssues([
    ['Equipment Name', 'UPN Tag', 'Cx Step', 'Category', 'Root Cause'],
    ['Pump A', '1001', 'Functional Test', 'COR', 'Design'],
    ['Pump A', '1001', 'Functional Test', 'NCR-D', 'Install'],
    ['Panel B', '2002', 'Startup', 'WI', 'Design'],
  ]), { byId: new Map([['pump a', 'Pump'], ['panel b', 'Panel']]) })

  const rows = ctx.stepCaughtRows(model.issueRows, model.equipment)
  const functional = rows.find((row) => row.key === 'Functional Test')

  assert.equal(JSON.stringify(rows.map((row) => row.key)), JSON.stringify(['Functional Test', 'Startup']))
  assert.equal(functional.issueCount, 2)
  assert.equal(functional.categoryCounts.COR, 1)
  assert.equal(functional.categoryCounts['NCR-D'], 1)
  assert.equal(JSON.stringify(functional.segments.map((segment) => [segment.category, segment.count, segment.ratio])), JSON.stringify([['COR', 1, 50], ['NCR-D', 1, 50]]))
})

test('matrix expanded score set toggles without losing other expanded sections', () => {
  const ctx = loadPipeline()

  assert.equal(JSON.stringify(ctx.toggleExpandedScore([1, 3], 2)), JSON.stringify([1, 2, 3]))
  assert.equal(JSON.stringify(ctx.toggleExpandedScore([1, 2, 3], 1)), JSON.stringify([2, 3]))
})

test('risk breakdown groups include nested equipment sorted by issue count', () => {
  const ctx = loadPipeline()
  const rows = [
    { name: 'Panel A', systemName: 'Power', discipline: 'Electrical', building: 'B1', milestone: 'M1', classification: 'Panel', score: 2, issueCount: 2, hasIssues: true, present: { QAQC: true, DV: true, EHS: false }, categories: ['QAQC', 'DV'] },
    { name: 'Panel B', systemName: 'Power', discipline: 'Electrical', building: 'B1', milestone: 'M1', classification: 'Panel', score: 3, issueCount: 5, hasIssues: true, present: { QAQC: true, DV: true, EHS: true }, categories: ['QAQC', 'DV', 'EHS'] },
    { name: 'Pump A', systemName: 'Water', discipline: 'Mechanical', building: 'B2', milestone: 'M2', classification: 'Pump', score: 1, issueCount: 1, hasIssues: true, present: { QAQC: true, DV: false, EHS: false }, categories: ['QAQC'] },
  ]

  const groups = ctx.riskBreakdownGroups(rows, 'Classification', 'hot')
  const panel = groups.find((group) => group.key === 'Panel')

  assert.equal(panel.n, 2)
  assert.equal(panel.issues, 7)
  assert.equal(JSON.stringify(panel.equipment.map((row) => row.name)), JSON.stringify(['Panel B', 'Panel A']))
  assert.equal(JSON.stringify(panel.equipment.map((row) => row.issueCount)), JSON.stringify([5, 2]))
})

test('coverage mix segments keep every score tier available for drilldown labels', () => {
  const ctx = loadPipeline()
  const segments = ctx.coverageMixSegments({ 0: 2, 1: 0, 2: 3, 3: 5 }, 10)

  assert.equal(JSON.stringify(segments.map((segment) => segment.label)), JSON.stringify(['0/3', '1/3', '2/3', '3/3']))
  assert.equal(JSON.stringify(segments.map((segment) => segment.count)), JSON.stringify([2, 0, 3, 5]))
  assert.equal(JSON.stringify(segments.map((segment) => segment.pctLabel)), JSON.stringify(['20%', '0%', '30%', '50%']))
})

test('chart metric sorting supports high-to-low and low-to-high directions', () => {
  const ctx = loadPipeline()
  const rows = [
    { key: 'Middle', issueCount: 4 },
    { key: 'Low', issueCount: 1 },
    { key: 'High', issueCount: 9 },
  ]

  assert.equal(JSON.stringify(ctx.sortRowsByMetric(rows, 'issueCount', 'desc').map((row) => row.key)), JSON.stringify(['High', 'Middle', 'Low']))
  assert.equal(JSON.stringify(ctx.sortRowsByMetric(rows, 'issueCount', 'asc').map((row) => row.key)), JSON.stringify(['Low', 'Middle', 'High']))
})
