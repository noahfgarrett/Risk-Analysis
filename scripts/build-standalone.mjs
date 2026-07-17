import { mkdirSync, readFileSync, readdirSync, rmSync, writeFileSync } from 'node:fs'
import { dirname, join } from 'node:path'
import { fileURLToPath } from 'node:url'

const root = dirname(dirname(fileURLToPath(import.meta.url)))
const dist = join(root, 'dist')
const pkg = JSON.parse(readFileSync(join(root, 'package.json'), 'utf8'))
const outputName = `Risk-Analysis-v${pkg.version}.html`
const replacements = [
  ['<script id="xlsx-runtime" src="vendor/xlsx.bundle.js"></script>', 'vendor/xlsx.bundle.js', ' id="xlsx-runtime"'],
  ['<script src="vendor/chart.umd.min.js"></script>', 'vendor/chart.umd.min.js', ''],
  ['<script src="vendor/html2canvas.min.js"></script>', 'vendor/html2canvas.min.js', ''],
]

let html = readFileSync(join(root, 'index.html'), 'utf8')

for (const [tag, assetPath, attributes] of replacements) {
  const source = readFileSync(join(root, assetPath), 'utf8').replace(/<\/script/gi, '<\\/script')
  html = html.replace(tag, () => `<script${attributes}>\n${source}\n</script>`)
}

mkdirSync(dist, { recursive: true })
for (const file of readdirSync(dist)) {
  if (/^Risk-Analysis.*\.html(\.gz)?$/.test(file)) rmSync(join(dist, file))
}
writeFileSync(join(dist, outputName), html)

console.log(`Built dist/${outputName}`)
