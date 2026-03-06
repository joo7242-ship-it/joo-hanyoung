import { Hono } from 'hono'

const app = new Hono()

// ── 헬스체크 ─────────────────────────────────────────────
app.get('/api/health', (c) => c.json({ status: 'ok', ts: Date.now() }))

// ── 미리보기 API ──────────────────────────────────────────
// Worker가 /docs/ 정적 파일을 self-fetch → ZIP 파싱 → HTML 변환
app.get('/api/preview/:std/:filename', async (c) => {
  const std = c.req.param('std')
  const filename = c.req.param('filename')

  try {
    const baseUrl = new URL(c.req.url)
    const fileUrl = `${baseUrl.origin}/docs/${std}/${filename}`
    const res = await fetch(fileUrl)
    if (!res.ok) return c.json({ error: '파일을 찾을 수 없습니다' }, 404)

    const buf = await res.arrayBuffer()
    const fname = decodeURIComponent(filename)
    const ext = fname.split('.').pop()?.toLowerCase()

    if (ext === 'docx') {
      const html = await parseDocx(buf)
      return c.json({ html, type: 'docx' })
    } else if (ext === 'xlsx') {
      const result = await parseXlsx(buf)
      return c.json({ sheets: result, type: 'xlsx' })
    } else {
      return c.json({ error: '지원하지 않는 형식입니다' }, 400)
    }
  } catch (e: any) {
    return c.json({ error: e.message || '파싱 오류' }, 500)
  }
})

// ── 메인 페이지 ───────────────────────────────────────────
app.get('/', (c) => {
  return c.html(renderHTML())
})

app.notFound((c) => {
  return c.html(renderHTML(), 200)
})

// ═══════════════════════════════════════════════════════
// ZIP 파서 (순수 JS / Cloudflare Workers 호환)
// ═══════════════════════════════════════════════════════
async function readZip(buf: ArrayBuffer): Promise<Map<string, Uint8Array>> {
  const data = new Uint8Array(buf)
  const files = new Map<string, Uint8Array>()

  // End of Central Directory 찾기
  let eocdOffset = -1
  for (let i = data.length - 22; i >= 0; i--) {
    if (data[i] === 0x50 && data[i+1] === 0x4b && data[i+2] === 0x05 && data[i+3] === 0x06) {
      eocdOffset = i; break
    }
  }
  if (eocdOffset < 0) throw new Error('ZIP EOCD not found')

  const view = new DataView(buf)
  const cdOffset = view.getUint32(eocdOffset + 16, true)
  const cdCount  = view.getUint16(eocdOffset + 8,  true)

  let pos = cdOffset
  for (let i = 0; i < cdCount; i++) {
    if (view.getUint32(pos, true) !== 0x02014b50) break
    const compMethod   = view.getUint16(pos + 10, true)
    const compSize     = view.getUint32(pos + 20, true)
    const uncompSize   = view.getUint32(pos + 24, true)
    const fnLen        = view.getUint16(pos + 28, true)
    const extraLen     = view.getUint16(pos + 30, true)
    const commentLen   = view.getUint16(pos + 32, true)
    const localOffset  = view.getUint32(pos + 42, true)
    const fname = new TextDecoder('utf-8').decode(data.slice(pos + 46, pos + 46 + fnLen))
    pos += 46 + fnLen + extraLen + commentLen

    // Local file header
    const lh = localOffset
    const lfnLen  = view.getUint16(lh + 26, true)
    const lextraLen = view.getUint16(lh + 28, true)
    const dataStart = lh + 30 + lfnLen + lextraLen
    const compData = data.slice(dataStart, dataStart + compSize)

    if (compMethod === 0) {
      files.set(fname, compData)
    } else if (compMethod === 8) {
      try {
        const ds = new DecompressionStream('deflate-raw')
        const writer = ds.writable.getWriter()
        writer.write(compData)
        writer.close()
        const reader = ds.readable.getReader()
        const chunks: Uint8Array[] = []
        while (true) {
          const { done, value } = await reader.read()
          if (done) break
          chunks.push(value)
        }
        const out = new Uint8Array(uncompSize)
        let offset = 0
        for (const chunk of chunks) { out.set(chunk, offset); offset += chunk.length }
        files.set(fname, out)
      } catch { /* skip unreadable entries */ }
    }
  }
  return files
}

// ═══════════════════════════════════════════════════════
// DOCX 파서
// ═══════════════════════════════════════════════════════
async function parseDocx(buf: ArrayBuffer): Promise<string> {
  const zip = await readZip(buf)
  const docXml = zip.get('word/document.xml')
  if (!docXml) throw new Error('word/document.xml not found')

  const xml = new TextDecoder('utf-8').decode(docXml)

  // 관계 네임스페이스 제거 후 파싱
  const clean = xml
    .replace(/<\?xml[^>]*>/g, '')
    .replace(/xmlns[^"]*"[^"]*"/g, '')
    .replace(/mc:AlternateContent>[\s\S]*?<\/mc:AlternateContent>/g, '')

  const html = convertDocxXmlToHtml(clean)
  return html
}

function convertDocxXmlToHtml(xml: string): string {
  let html = '<div class="doc-content">'

  // 단락 추출
  const paraReg = /<w:p[ >]([\s\S]*?)<\/w:p>/g
  let match
  while ((match = paraReg.exec(xml)) !== null) {
    const para = match[1]

    // 스타일 확인
    const styleMatch = para.match(/<w:pStyle w:val="([^"]+)"/)
    const style = styleMatch ? styleMatch[1] : 'Normal'

    // 텍스트 추출 (w:t 태그)
    const texts: string[] = []
    const runReg = /<w:r[ >]([\s\S]*?)<\/w:r>/g
    let rMatch
    while ((rMatch = runReg.exec(para)) !== null) {
      const run = rMatch[1]
      // 볼드/이탤릭 체크
      const isBold   = /<w:b(?:\s+w:val="(?!0)[^"]*"|\s*\/>|\s*>)/.test(run) || /<w:b\/>/.test(run)
      const isItalic = /<w:i(?:\s+w:val="(?!0)[^"]*"|\s*\/>|\s*>)/.test(run) || /<w:i\/>/.test(run)
      const tMatch = run.match(/<w:t[^>]*>([\s\S]*?)<\/w:t>/)
      if (tMatch) {
        let t = esc(tMatch[1])
        if (isBold)   t = `<strong>${t}</strong>`
        if (isItalic) t = `<em>${t}</em>`
        texts.push(t)
      }
    }
    const text = texts.join('')
    if (!text.trim()) { html += '<p class="doc-empty">&nbsp;</p>'; continue }

    // 스타일별 태그
    const s = style.toLowerCase()
    if      (s.includes('heading 1') || s === 'heading1' || s.includes('1'))
      html += `<h1 class="doc-h1">${text}</h1>`
    else if (s.includes('heading 2') || s === 'heading2' || s.includes('2') && s.includes('head'))
      html += `<h2 class="doc-h2">${text}</h2>`
    else if (s.includes('heading 3') || s === 'heading3')
      html += `<h3 class="doc-h3">${text}</h3>`
    else if (s.includes('listbullet') || s.includes('list bullet') || s.includes('listparagraph'))
      html += `<li class="doc-li">${text}</li>`
    else
      html += `<p class="doc-p">${text}</p>`
  }

  // 표 추출
  const tableReg = /<w:tbl[ >]([\s\S]*?)<\/w:tbl>/g
  while ((match = tableReg.exec(xml)) !== null) {
    html += convertTableToHtml(match[1])
  }

  html += '</div>'
  return html
}

function convertTableToHtml(tblXml: string): string {
  let tbl = '<table class="doc-table"><tbody>'
  const rowReg = /<w:tr[ >]([\s\S]*?)<\/w:tr>/g
  let rMatch
  while ((rMatch = rowReg.exec(tblXml)) !== null) {
    tbl += '<tr>'
    const cellReg = /<w:tc[ >]([\s\S]*?)<\/w:tc>/g
    let cMatch
    while ((cMatch = cellReg.exec(rMatch[1])) !== null) {
      const cell = cMatch[1]
      const spanMatch = cell.match(/w:gridSpan w:val="(\d+)"/)
      const colspan = spanMatch ? ` colspan="${spanMatch[1]}"` : ''
      const texts: string[] = []
      const tReg = /<w:t[^>]*>([\s\S]*?)<\/w:t>/g
      let tMatch
      while ((tMatch = tReg.exec(cell)) !== null) texts.push(esc(tMatch[1]))
      tbl += `<td${colspan} class="doc-td">${texts.join('')}</td>`
    }
    tbl += '</tr>'
  }
  tbl += '</tbody></table>'
  return tbl
}

// ═══════════════════════════════════════════════════════
// XLSX 파서
// ═══════════════════════════════════════════════════════
async function parseXlsx(buf: ArrayBuffer): Promise<{ name: string; html: string }[]> {
  const zip = await readZip(buf)

  // sharedStrings
  const ssData = zip.get('xl/sharedStrings.xml')
  const sharedStrings: string[] = []
  if (ssData) {
    const ssXml = new TextDecoder('utf-8').decode(ssData)
    const siReg = /<si>([\s\S]*?)<\/si>/g
    let m
    while ((m = siReg.exec(ssXml)) !== null) {
      const tReg = /<t[^>]*>([^<]*)<\/t>/g
      let tm; const parts: string[] = []
      while ((tm = tReg.exec(m[1])) !== null) parts.push(tm[1])
      sharedStrings.push(parts.join(''))
    }
  }

  // workbook.xml → 시트 목록
  const wbData = zip.get('xl/workbook.xml')
  const sheetNames: { name: string; id: string }[] = []
  if (wbData) {
    const wbXml = new TextDecoder('utf-8').decode(wbData)
    const shReg = /<sheet[^>]+name="([^"]+)"[^>]+r:id="([^"]+)"/g
    let m
    while ((m = shReg.exec(wbXml)) !== null) sheetNames.push({ name: m[1], id: m[2] })
  }

  // workbook.xml.rels → sheet 파일 경로
  const relsData = zip.get('xl/_rels/workbook.xml.rels')
  const relMap = new Map<string, string>()
  if (relsData) {
    const relsXml = new TextDecoder('utf-8').decode(relsData)
    const relReg = /<Relationship Id="([^"]+)"[^>]+Target="([^"]+)"/g
    let m
    while ((m = relReg.exec(relsXml)) !== null) relMap.set(m[1], m[2])
  }

  const results: { name: string; html: string }[] = []

  for (const sh of sheetNames.length > 0 ? sheetNames : [{ name: 'Sheet1', id: 'rId1' }]) {
    const target = relMap.get(sh.id) || `worksheets/sheet${results.length + 1}.xml`
    const path = target.startsWith('xl/') ? target : `xl/${target}`
    const shData = zip.get(path)
    if (!shData) continue

    const shXml = new TextDecoder('utf-8').decode(shData)
    const html = convertSheetToHtml(shXml, sharedStrings)
    results.push({ name: sh.name, html })
  }

  if (results.length === 0) {
    // fallback: 첫 번째 sheet 파일 시도
    for (const [k, v] of zip) {
      if (k.startsWith('xl/worksheets/sheet') && k.endsWith('.xml')) {
        const html = convertSheetToHtml(new TextDecoder('utf-8').decode(v), sharedStrings)
        results.push({ name: 'Sheet1', html })
        break
      }
    }
  }

  return results
}

function convertSheetToHtml(xml: string, ss: string[]): string {
  // 행 추출
  const rows: string[][] = []
  const rowReg = /<row[^>]*>([\s\S]*?)<\/row>/g
  let rMatch
  while ((rMatch = rowReg.exec(xml)) !== null) {
    const rowXml = rMatch[1]
    const cells: string[] = []
    const cellReg = /<c ([^>]*)>([\s\S]*?)<\/c>/g
    let cMatch
    while ((cMatch = cellReg.exec(rowXml)) !== null) {
      const attrs = cMatch[0]
      const cellContent = cMatch[2]
      const tAttr = attrs.match(/\bt="([^"]+)"/)
      const cellType = tAttr ? tAttr[1] : ''
      const vMatch = cellContent.match(/<v>([^<]*)<\/v>/)
      const isMatch = cellContent.match(/<is>([\s\S]*?)<\/is>/)

      let val = ''
      if (isMatch) {
        const tReg = /<t[^>]*>([^<]*)<\/t>/g
        let tm; const parts: string[] = []
        while ((tm = tReg.exec(isMatch[1])) !== null) parts.push(tm[1])
        val = parts.join('')
      } else if (vMatch) {
        if (cellType === 's') {
          val = ss[parseInt(vMatch[1])] ?? ''
        } else if (cellType === 'b') {
          val = vMatch[1] === '1' ? 'TRUE' : 'FALSE'
        } else {
          val = vMatch[1]
        }
      }
      cells.push(esc(val))
    }
    if (cells.some(c => c.trim())) rows.push(cells)
  }

  if (rows.length === 0) return '<p style="color:#aaa;text-align:center;padding:20px">데이터 없음</p>'

  // 최대 컬럼 수
  const maxCols = Math.max(...rows.map(r => r.length))

  // 첫 번째 비어있지 않은 행 = 헤더
  let headerIdx = 0
  for (let i = 0; i < rows.length; i++) {
    if (rows[i].some(c => c.trim())) { headerIdx = i; break }
  }

  let html = '<table class="xl-table"><thead><tr>'
  const headerRow = rows[headerIdx]
  for (let c = 0; c < maxCols; c++) {
    html += `<th class="xl-th">${headerRow[c] ?? ''}</th>`
  }
  html += '</tr></thead><tbody>'

  for (let i = headerIdx + 1; i < rows.length; i++) {
    html += '<tr>'
    for (let c = 0; c < maxCols; c++) {
      html += `<td class="xl-td">${rows[i][c] ?? ''}</td>`
    }
    html += '</tr>'
  }
  html += '</tbody></table>'
  return html
}

function esc(s: string): string {
  return s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')
}

// ═══════════════════════════════════════════════════════
// HTML 렌더러
// ═══════════════════════════════════════════════════════
function renderHTML(): string {
  return `<!DOCTYPE html>
<html lang="ko">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width,initial-scale=1.0">
<title>(주)한영피엔에스 IMS 문서 미리보기 시스템</title>
<link rel="icon" href="data:image/svg+xml,<svg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 32 32'><rect width='32' height='32' rx='6' fill='%231F3864'/><text x='16' y='22' font-size='18' text-anchor='middle' fill='white' font-family='Arial'>P</text></svg>">
<style>
*{box-sizing:border-box;margin:0;padding:0}
:root{--brand:#1F3864;--brand-mid:#2E75B6;--brand-light:#D6E4F0;--sw:260px;--fw:290px;--hh:60px}
body{font-family:'Malgun Gothic','맑은 고딕',sans-serif;background:#f0f2f5;color:#222}
/* ── 헤더 ── */
.header{position:fixed;top:0;left:0;right:0;height:var(--hh);background:var(--brand);color:#fff;display:flex;align-items:center;padding:0 16px;z-index:100;box-shadow:0 2px 8px rgba(0,0,0,.3);gap:12px}
.header-logo{font-size:16px;font-weight:700;white-space:nowrap;line-height:1.3}
.header-logo span{font-size:10px;opacity:.7;display:block;font-weight:400}
.header-search{flex:1;max-width:420px;position:relative}
.header-search input{width:100%;padding:7px 14px 7px 34px;border:none;border-radius:18px;font-size:13px;background:rgba(255,255,255,.18);color:#fff;outline:none;font-family:inherit}
.header-search input::placeholder{color:rgba(255,255,255,.55)}
.header-search input:focus{background:rgba(255,255,255,.28)}
.si{position:absolute;left:11px;top:50%;transform:translateY(-50%);opacity:.7;font-size:14px}
.sr{position:absolute;top:40px;left:0;right:0;background:#fff;border-radius:8px;box-shadow:0 8px 24px rgba(0,0,0,.18);max-height:320px;overflow-y:auto;z-index:200;display:none;border:1px solid #e0e0e0}
.sr.show{display:block}
.sri{padding:8px 12px;cursor:pointer;display:flex;align-items:center;gap:8px;border-bottom:1px solid #f3f3f3;color:#222;font-size:12px}
.sri:hover{background:var(--brand-light)}
.sr0{padding:12px 14px;color:#aaa;font-size:13px;text-align:center}
.st{font-size:10px;padding:2px 5px;border-radius:4px;color:#fff;white-space:nowrap}
.hstats{font-size:11px;opacity:.85;white-space:nowrap;display:flex;gap:10px}
.sitem{text-align:center}.sitem strong{display:block;font-size:16px;font-weight:800}
/* ── 레이아웃 ── */
.layout{display:flex;margin-top:var(--hh);height:calc(100vh - var(--hh));overflow:hidden}
/* ── 사이드바 ── */
.sidebar{width:var(--sw);background:#fff;overflow-y:auto;border-right:1px solid #e0e0e0;flex-shrink:0}
.sb-title{padding:10px 14px 6px;font-size:10px;color:#999;text-transform:uppercase;font-weight:700;letter-spacing:.6px;border-bottom:1px solid #f0f0f0}
.std-card{display:flex;align-items:center;gap:8px;padding:9px 12px;cursor:pointer;transition:.12s;border-left:3px solid transparent}
.std-card:hover{background:#f8f9fb}
.std-card.active{background:var(--brand-light);border-left-color:var(--brand)}
.std-icon{font-size:17px;width:24px;text-align:center}
.std-info{flex:1;min-width:0}
.std-name{font-size:11px;font-weight:700;white-space:nowrap;overflow:hidden;text-overflow:ellipsis}
.std-meta{font-size:10px;color:#999;margin-top:1px}
.std-badge{font-size:9px;padding:2px 5px;border-radius:4px;color:#fff;font-weight:700;white-space:nowrap}
/* ── 파일 패널 ── */
.fp{width:var(--fw);background:#fafafa;overflow-y:auto;border-right:1px solid #e0e0e0;flex-shrink:0;display:flex;flex-direction:column}
.fp-hdr{position:sticky;top:0;background:#fff;padding:10px 12px;border-bottom:1px solid #e8e8e8;z-index:10;flex-shrink:0}
.fp-title{font-size:13px;font-weight:700;color:var(--brand)}
.fp-count{font-size:10px;color:#999;margin-top:1px}
.fp-filter{display:flex;gap:3px;margin-top:7px;flex-wrap:wrap}
.fbtn{padding:3px 8px;border-radius:10px;border:1px solid #ddd;font-size:10px;cursor:pointer;background:#fff;color:#555;transition:.12s;font-family:inherit}
.fbtn.active{background:var(--brand);color:#fff;border-color:var(--brand)}
.fbtn:hover:not(.active){background:#f0f0f0}
.fl{padding:5px;flex:1}
.fi{display:flex;align-items:center;gap:6px;padding:6px 8px;border-radius:5px;cursor:pointer;margin-bottom:2px;transition:.12s}
.fi:hover{background:#eff2f7}
.fi.active{background:var(--brand-light);box-shadow:inset 0 0 0 1px var(--brand-mid)}
.fext{font-size:9px;padding:2px 5px;border-radius:3px;font-weight:800;color:#fff;min-width:32px;text-align:center;flex-shrink:0}
.fext.docx{background:#2E75B6}.fext.xlsx{background:#217346}
.fn{font-size:11px;flex:1;min-width:0}
.fn strong{display:block;font-weight:600;color:#222;white-space:nowrap;overflow:hidden;text-overflow:ellipsis}
.fn small{color:#aaa;font-size:10px}
.ftb{font-size:9px;padding:1px 5px;border-radius:3px;color:#fff;flex-shrink:0}
/* ── 미리보기 패널 ── */
.pp{flex:1;overflow:hidden;display:flex;flex-direction:column;min-width:0}
.ph{background:#fff;padding:10px 16px;border-bottom:1px solid #e8e8e8;display:none;align-items:center;gap:10px;flex-shrink:0}
.ph.show{display:flex}
.ph-fn{flex:1;min-width:0}
.ph-fn h3{font-size:13px;color:var(--brand);font-weight:700;white-space:nowrap;overflow:hidden;text-overflow:ellipsis}
.ph-fn p{font-size:10px;color:#999;margin-top:1px}
.btn-grp{display:flex;gap:6px;flex-shrink:0}
.bdl{padding:5px 12px;background:var(--brand);color:#fff;border:none;border-radius:5px;cursor:pointer;font-size:12px;display:flex;align-items:center;gap:4px;white-space:nowrap;text-decoration:none;font-family:inherit;transition:.12s}
.bdl:hover{background:#162a4d}
.bdl.green{background:#217346}.bdl.green:hover{background:#155230}
/* ── Excel 시트 탭 ── */
.sheet-tabs{display:flex;gap:2px;padding:6px 16px 0;background:#f8f8f8;border-bottom:1px solid #e0e0e0;flex-shrink:0;overflow-x:auto}
.stab{padding:5px 14px;font-size:12px;cursor:pointer;border-radius:4px 4px 0 0;border:1px solid transparent;background:#e8e8e8;color:#555;white-space:nowrap;font-family:inherit;transition:.12s}
.stab.active{background:#fff;border-color:#e0e0e0 #e0e0e0 #fff;color:var(--brand);font-weight:700}
.stab:hover:not(.active){background:#ddd}
/* ── 미리보기 본문 ── */
.pb{flex:1;overflow-y:auto;padding:20px;background:#fafafa}
/* ── 로딩 ── */
.loading{display:flex;align-items:center;justify-content:center;height:200px;flex-direction:column;gap:12px}
.spinner{width:32px;height:32px;border:3px solid #e8e8e8;border-top-color:var(--brand);border-radius:50%;animation:spin .7s linear infinite}
@keyframes spin{to{transform:rotate(360deg)}}
.lt{font-size:12px;color:#aaa}
/* ── 환영 화면 ── */
.welcome{display:flex;flex-direction:column;align-items:center;justify-content:center;height:100%;color:#bbb;text-align:center;padding:40px}
.welcome h2{font-size:20px;color:#ccc;margin-bottom:10px}
.welcome p{font-size:13px;line-height:1.9}
/* ── 토스트 ── */
.toast{position:fixed;bottom:24px;right:24px;background:#333;color:#fff;padding:10px 18px;border-radius:6px;font-size:13px;z-index:9999;opacity:0;transition:opacity .3s;pointer-events:none}
.toast.show{opacity:1}
/* ── DOCX 스타일 ── */
.doc-content{max-width:820px;margin:0 auto;background:#fff;padding:32px 40px;border-radius:8px;box-shadow:0 1px 6px rgba(0,0,0,.08);font-size:13px;line-height:1.8;color:#222}
.doc-h1{font-size:18px;font-weight:700;color:var(--brand);margin:20px 0 10px;padding-bottom:6px;border-bottom:2px solid var(--brand-light)}
.doc-h2{font-size:15px;font-weight:700;color:var(--brand-mid);margin:16px 0 8px}
.doc-h3{font-size:13px;font-weight:700;color:#444;margin:12px 0 6px}
.doc-p{margin:4px 0;color:#333}
.doc-empty{height:6px}
.doc-li{margin:3px 0 3px 20px;list-style:disc;color:#333}
.doc-table{width:100%;border-collapse:collapse;margin:12px 0;font-size:12px}
.doc-td{border:1px solid #ccc;padding:5px 8px;vertical-align:top}
/* ── XLSX 스타일 ── */
.xl-wrap{overflow-x:auto;padding:4px 0}
.xl-table{border-collapse:collapse;min-width:600px;font-size:12px;background:#fff}
.xl-th{background:#1F3864;color:#fff;padding:7px 10px;text-align:left;font-weight:700;white-space:nowrap;border:1px solid #16305a;font-size:11px}
.xl-td{padding:5px 10px;border:1px solid #e0e0e0;vertical-align:top;white-space:nowrap;max-width:300px;overflow:hidden;text-overflow:ellipsis}
tr:nth-child(even) .xl-td{background:#f8faff}
tr:hover .xl-td{background:#e8f0ff}
.empty-state{text-align:center;padding:32px 16px;color:#ccc;font-size:13px}
/* ── 에러 카드 ── */
.err-card{background:#fff5f5;border:1px solid #fcc;border-radius:8px;padding:24px;max-width:500px;margin:40px auto;text-align:center;color:#c00}
</style>
</head>
<body>
<header class="header">
  <div class="header-logo">(주)한영피엔에스<span>IMS 통합경영시스템 문서 미리보기</span></div>
  <div class="header-search">
    <span class="si">🔍</span>
    <input type="text" id="searchInput" placeholder="문서명 검색... (2자 이상)" autocomplete="off">
    <div class="sr" id="sr"></div>
  </div>
  <div class="hstats">
    <div class="sitem"><strong id="stTotal">-</strong>총 문서</div>
    <div class="sitem"><strong id="stWord">-</strong>Word</div>
    <div class="sitem"><strong id="stExcel">-</strong>Excel</div>
  </div>
</header>
<div class="layout">
  <aside class="sidebar">
    <div class="sb-title">📌 ISO 규격 선택</div>
    <div id="stdList"></div>
  </aside>
  <div class="fp" id="fp">
    <div class="welcome" style="height:100%">
      <div style="font-size:44px;margin-bottom:10px">📂</div>
      <h2>규격을 선택하세요</h2>
      <p>좌측에서 ISO 규격을 클릭하면<br>해당 문서 목록이 표시됩니다</p>
    </div>
  </div>
  <div class="pp" id="pp">
    <div class="ph" id="ph">
      <div class="ph-fn"><h3 id="phTitle">-</h3><p id="phMeta">-</p></div>
      <div class="btn-grp" id="btnGrp">
        <a id="btnDl" href="#" class="bdl">⬇ 다운로드</a>
      </div>
    </div>
    <div id="sheetTabsArea"></div>
    <div class="pb" id="pb">
      <div class="welcome">
        <div style="font-size:56px;margin-bottom:14px">📄</div>
        <h2>문서를 선택하세요</h2>
        <p>(주)한영피엔에스 IMS 통합경영시스템<br>
        좌측에서 규격 → 문서를 선택하면<br>내용이 미리보기로 표시됩니다.<br><br>
        <span style="color:#ddd;font-size:12px">144개 문서 · 7개 ISO 규격 + IMS 통합</span></p>
      </div>
    </div>
  </div>
</div>
<div class="toast" id="toast"></div>
<script>
// ── 상태 ──
let META = null;
let curStd = null, curFile = null, allFiles = [], activeFilter = 'all';
let xlSheets = [], xlActiveSheet = 0;
const previewCache = new Map();

// ── 유틸 ──
function toast(msg, ms=2200){
  const t=document.getElementById('toast');
  t.textContent=msg; t.classList.add('show');
  setTimeout(()=>t.classList.remove('show'),ms);
}
function esc(s){ return String(s).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;'); }
function setLoading(){
  document.getElementById('sheetTabsArea').innerHTML='';
  document.getElementById('pb').innerHTML='<div class="loading"><div class="spinner"></div><div class="lt">문서 로딩 중...</div></div>';
}

// ── 초기화 ──
async function init(){
  try{
    const r = await fetch('/static/docs-meta.json');
    META = await r.json();
    renderStats(META.stats);
    renderStdList(META.standards);
  }catch(e){ console.error('메타데이터 로드 실패',e); }
}

function renderStats(s){
  document.getElementById('stTotal').textContent = s.total;
  document.getElementById('stWord').textContent = s.word;
  document.getElementById('stExcel').textContent = s.excel;
}

function renderStdList(stds){
  const el = document.getElementById('stdList');
  el.innerHTML = stds.map(s=>\`
    <div class="std-card" data-std="\${s.key}" id="std-\${s.key}">
      <div class="std-icon">\${s.icon}</div>
      <div class="std-info">
        <div class="std-name">\${s.name}</div>
        <div class="std-meta">문서 \${s.count}개</div>
      </div>
      <div class="std-badge" style="background:\${s.color}">\${s.badge}</div>
    </div>\`).join('');
  el.addEventListener('click', e=>{
    const c=e.target.closest('.std-card');
    if(c) selectStd(c.dataset.std);
  });
}

// ── 규격 선택 ──
function selectStd(key){
  if(curStd===key) return;
  curStd=key; curFile=null; activeFilter='all';
  document.querySelectorAll('.std-card').forEach(e=>e.classList.remove('active'));
  document.getElementById('std-'+key)?.classList.add('active');
  document.getElementById('ph').classList.remove('show');
  document.getElementById('sheetTabsArea').innerHTML='';
  document.getElementById('pb').innerHTML=\`<div class="welcome"><div style="font-size:44px;margin-bottom:12px">👆</div><h2>문서를 선택하세요</h2><p>목록에서 문서를 클릭하면 미리보기가 표시됩니다</p></div>\`;
  const std = META.standards.find(s=>s.key===key);
  if(!std) return;
  allFiles = std.files;
  renderFilePanel(allFiles, key);
}

function renderFilePanel(files, key){
  const fp = document.getElementById('fp');
  const types = [...new Set(files.map(f=>f.type))];
  const filtered = activeFilter==='all' ? files : files.filter(f=>f.type===activeFilter);
  fp.innerHTML = \`
    <div class="fp-hdr">
      <div class="fp-title">\${key}</div>
      <div class="fp-count">전체 \${files.length}개 · 표시 \${filtered.length}개</div>
      <div class="fp-filter">
        <button class="fbtn \${activeFilter==='all'?'active':''}" data-filter="all">전체</button>
        \${types.map(t=>\`<button class="fbtn \${activeFilter===t?'active':''}" data-filter="\${t}">\${t}</button>\`).join('')}
      </div>
    </div>
    <div class="fl" id="fl">
      \${filtered.length===0
        ? '<div class="empty-state">해당 유형의 문서가 없습니다</div>'
        : filtered.map(f=>\`
        <div class="fi" data-file="\${encodeURIComponent(f.name)}" data-std="\${key}">
          <div class="fext \${f.ext.toLowerCase()}">\${f.ext}</div>
          <div class="fn"><strong title="\${esc(f.display)}">\${esc(f.display)}</strong><small>\${f.size} KB</small></div>
          <div class="ftb" style="background:\${f.color}">\${f.type}</div>
        </div>\`).join('')}
    </div>\`;
  document.getElementById('fl')?.addEventListener('click', e=>{
    const it=e.target.closest('.fi');
    if(it) showDoc(it.dataset.std, decodeURIComponent(it.dataset.file), it);
  });
  fp.querySelector('.fp-filter')?.addEventListener('click', e=>{
    const b=e.target.closest('.fbtn');
    if(b){ activeFilter=b.dataset.filter; renderFilePanel(allFiles,key); }
  });
}

// ── 미리보기 ──
async function showDoc(std, fname, itemEl){
  if(curFile===fname) return;
  curFile=fname;
  document.querySelectorAll('.fi').forEach(e=>e.classList.remove('active'));
  itemEl?.classList.add('active');

  const ext = fname.split('.').pop().toUpperCase();
  const display = fname.replace(/\\.(docx|xlsx)$/i,'');
  const dlUrl = '/docs/'+std+'/'+encodeURIComponent(fname);

  // 헤더 업데이트
  document.getElementById('phTitle').textContent = display;
  document.getElementById('phMeta').textContent = std+' · '+ext+' 문서';
  const btnDl = document.getElementById('btnDl');
  btnDl.href = dlUrl;
  btnDl.className = ext==='XLSX' ? 'bdl green' : 'bdl';
  btnDl.textContent = '⬇ 다운로드';
  document.getElementById('ph').classList.add('show');

  // 로딩
  setLoading();

  // 캐시 확인
  const cacheKey = std+'/'+fname;
  if(previewCache.has(cacheKey)){
    renderPreview(previewCache.get(cacheKey), ext, fname, std);
    return;
  }

  try{
    const apiUrl = '/api/preview/'+std+'/'+encodeURIComponent(fname);
    const r = await fetch(apiUrl);
    if(!r.ok){
      const err = await r.json().catch(()=>({error:'요청 실패'}));
      showError(err.error || '미리보기를 불러올 수 없습니다');
      return;
    }
    const data = await r.json();
    previewCache.set(cacheKey, data);
    renderPreview(data, ext, fname, std);
  }catch(e){
    showError('네트워크 오류: '+e.message);
  }
}

function renderPreview(data, ext, fname, std){
  document.getElementById('sheetTabsArea').innerHTML='';

  if(data.type==='docx'){
    document.getElementById('pb').innerHTML = data.html;
  } else if(data.type==='xlsx'){
    xlSheets = data.sheets;
    xlActiveSheet = 0;
    renderSheetTabs();
    renderSheet(0);
  }
}

function renderSheetTabs(){
  if(xlSheets.length<=1){
    document.getElementById('sheetTabsArea').innerHTML='';
    return;
  }
  const tabs = xlSheets.map((s,i)=>\`<button class="stab \${i===xlActiveSheet?'active':''}" data-idx="\${i}">\${esc(s.name)}</button>\`).join('');
  document.getElementById('sheetTabsArea').innerHTML='<div class="sheet-tabs">'+tabs+'</div>';
  document.getElementById('sheetTabsArea').addEventListener('click', e=>{
    const b=e.target.closest('.stab');
    if(b){ xlActiveSheet=parseInt(b.dataset.idx); renderSheetTabs(); renderSheet(xlActiveSheet); }
  });
}

function renderSheet(idx){
  const sheet = xlSheets[idx];
  if(!sheet){ document.getElementById('pb').innerHTML='<p style="color:#aaa;padding:20px">시트 없음</p>'; return; }
  document.getElementById('pb').innerHTML='<div class="xl-wrap">'+sheet.html+'</div>';
}

function showError(msg){
  document.getElementById('sheetTabsArea').innerHTML='';
  document.getElementById('pb').innerHTML=\`<div class="err-card"><div style="font-size:32px;margin-bottom:10px">⚠️</div><strong>미리보기 오류</strong><p style="margin-top:8px;font-size:12px;color:#a00">\${esc(msg)}</p></div>\`;
}

// ── 검색 ──
let st;
const sinput=document.getElementById('searchInput');
const srbox=document.getElementById('sr');
sinput.addEventListener('input',function(){
  clearTimeout(st);
  const q=this.value.trim();
  if(q.length<2){ srbox.classList.remove('show'); return; }
  st=setTimeout(()=>doSearch(q),220);
});
document.addEventListener('click', e=>{
  if(!e.target.closest('.header-search')) srbox.classList.remove('show');
});
sinput.addEventListener('focus',function(){
  if(this.value.trim().length>=2) srbox.classList.add('show');
});
function doSearch(q){
  if(!META) return;
  const ql=q.toLowerCase();
  const results=[];
  for(const std of META.standards){
    for(const f of std.files){
      if(f.name.toLowerCase().includes(ql)||f.display.toLowerCase().includes(ql)||std.name.toLowerCase().includes(ql)){
        results.push({...f, stdKey:std.key});
        if(results.length>=40) break;
      }
    }
    if(results.length>=40) break;
  }
  if(results.length===0){
    srbox.innerHTML='<div class="sr0">검색 결과가 없습니다</div>';
  } else {
    srbox.innerHTML=results.map(r=>\`
      <div class="sri" data-std="\${r.stdKey}" data-file="\${encodeURIComponent(r.name)}">
        <span class="st" style="background:\${r.color}">\${r.type}</span>
        <span style="flex:1;overflow:hidden;text-overflow:ellipsis;white-space:nowrap">\${esc(r.display)}</span>
        <span style="font-size:10px;color:#bbb;white-space:nowrap">\${r.stdKey}</span>
      </div>\`).join('');
    srbox.addEventListener('click', function handler(e){
      const si=e.target.closest('.sri');
      if(!si) return;
      const std=si.dataset.std; const file=decodeURIComponent(si.dataset.file);
      srbox.classList.remove('show'); sinput.value='';
      selectStd(std);
      setTimeout(()=>{
        const items=document.querySelectorAll('.fi');
        for(const it of items){
          if(decodeURIComponent(it.dataset.file)===file){ it.click(); it.scrollIntoView({block:'nearest',behavior:'smooth'}); break; }
        }
      },300);
      srbox.removeEventListener('click',handler);
    });
  }
  srbox.classList.add('show');
}

init();
</script>
</body>
</html>`
}

export default app
