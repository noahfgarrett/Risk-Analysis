import { mkdirSync, readFileSync, writeFileSync } from 'node:fs'
import { dirname, join } from 'node:path'
import { fileURLToPath } from 'node:url'

const root = dirname(dirname(fileURLToPath(import.meta.url)))
const dist = join(root, 'dist')
const replacements = [
  ['<script src="vendor/xlsx.bundle.js"></script>', 'vendor/xlsx.bundle.js'],
  ['<script src="vendor/chart.umd.min.js"></script>', 'vendor/chart.umd.min.js'],
  ['<script src="vendor/html2canvas.min.js"></script>', 'vendor/html2canvas.min.js'],
]

let html = readFileSync(join(root, 'index.html'), 'utf8')

for (const [tag, assetPath] of replacements) {
  const source = readFileSync(join(root, assetPath), 'utf8').replace(/<\/script/gi, '<\\/script')
  html = html.replace(tag, () => `<script>\n${source}\n</script>`)
}

mkdirSync(dist, { recursive: true })
writeFileSync(join(dist, 'Risk-Analysis.html'), html)

console.log('Built dist/Risk-Analysis.html')
