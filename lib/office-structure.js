/**
 * Lightweight, image-free structure scanning for Office documents and PDFs.
 *
 * Both viewers in this bundle previously had to guess a document's shape from
 * converted text, which cannot reveal images at all, or by rendering every page,
 * which spends vision tokens on pages that hold only text. This module reads the
 * container instead:
 *
 *   - OOXML (docx/xlsx/pptx) is a ZIP; sections, tables, pictures, headings and
 *     sheet/slide facts come straight from the part XML, so no converter runs.
 *   - Pagination and per-page pictures need the real layout engine, so Word
 *     documents are converted to PDF by the official Office service and the
 *     PDF's page tree is parsed. The PDF is never turned into an image, so a
 *     scan costs no vision tokens.
 *   - PDFs are parsed directly.
 *
 * Node built-ins only: no dependency, no shell, no writes.
 *
 * @module dsh-better-chat/office-structure
 */

import { inflateRawSync, inflateSync } from 'node:zlib'

/** Inflated package bytes a single scan accepts before refusing. */
const MAX_INFLATED_BYTES = 96 * 1024 * 1024

/* ------------------------------------------------------------------ *
 * ZIP (OOXML container)
 * ------------------------------------------------------------------ */

const EOCD_SIGNATURE = 0x06054b50
const CENTRAL_SIGNATURE = 0x02014b50
const LOCAL_SIGNATURE = 0x04034b50

/**
 * Open an OOXML package lazily: entry names come from the central directory and
 * one part inflates only when read, so embedded media never cost memory.
 * @param {Buffer} buf - complete package bytes.
 */
export function openZip(buf) {
  const eocd = findEocd(buf)
  if (eocd < 0) throw new Error('not a ZIP package (no end-of-central-directory record)')
  const count = buf.readUInt16LE(eocd + 10)
  const cdSize = buf.readUInt32LE(eocd + 12)
  const cdOffset = buf.readUInt32LE(eocd + 16)
  if (count === 0xffff || cdOffset === 0xffffffff || cdSize === 0xffffffff) {
    throw new Error('ZIP64 packages are not supported')
  }
  const entries = new Map()
  let p = cdOffset
  let inflatedTotal = 0
  for (let i = 0; i < count; i += 1) {
    if (p + 46 > buf.length || buf.readUInt32LE(p) !== CENTRAL_SIGNATURE) {
      throw new Error('damaged ZIP central directory')
    }
    const method = buf.readUInt16LE(p + 10)
    const csize = buf.readUInt32LE(p + 20)
    const usize = buf.readUInt32LE(p + 24)
    const nameLen = buf.readUInt16LE(p + 28)
    const extraLen = buf.readUInt16LE(p + 30)
    const commentLen = buf.readUInt16LE(p + 32)
    const local = buf.readUInt32LE(p + 42)
    const name = buf.toString('utf8', p + 46, p + 46 + nameLen)
    if (csize === 0xffffffff || usize === 0xffffffff || local === 0xffffffff) {
      throw new Error('ZIP64 entries are not supported')
    }
    if (!name.endsWith('/')) entries.set(name, { method, csize, usize, local })
    p += 46 + nameLen + extraLen + commentLen
  }

  function read(name) {
    const entry = entries.get(name)
    if (entry === undefined) return undefined
    const at = entry.local
    if (at + 30 > buf.length || buf.readUInt32LE(at) !== LOCAL_SIGNATURE) {
      throw new Error('damaged ZIP local header for ' + name)
    }
    const nameLen = buf.readUInt16LE(at + 26)
    const extraLen = buf.readUInt16LE(at + 28)
    const start = at + 30 + nameLen + extraLen
    const data = buf.subarray(start, start + entry.csize)
    let out
    if (entry.method === 0) out = Buffer.from(data)
    else if (entry.method === 8) out = inflateRawSync(data, entry.usize > 0 ? { maxOutputLength: entry.usize } : undefined)
    else throw new Error('unsupported ZIP compression method ' + entry.method + ' for ' + name)
    inflatedTotal += out.length
    if (inflatedTotal > MAX_INFLATED_BYTES) throw new Error('package expands beyond the scan limit')
    return out
  }

  return {
    names: () => [...entries.keys()],
    has: (name) => entries.has(name),
    read,
    text(name) {
      const bytes = read(name)
      return bytes === undefined ? undefined : bytes.toString('utf8')
    },
  }
}

function findEocd(buf) {
  const floor = Math.max(0, buf.length - 0xffff - 22)
  for (let i = buf.length - 22; i >= floor; i -= 1) {
    if (buf.readUInt32LE(i) === EOCD_SIGNATURE) return i
  }
  return -1
}

/* ------------------------------------------------------------------ *
 * XML tokenizer and helpers
 * ------------------------------------------------------------------ */

/**
 * Yield the flat tag/text stream of an XML part. Attribute values may contain
 * '>', so a tag ends at the first '>' outside quotes.
 * @param {string} src - XML source.
 */
export function* xmlEvents(src) {
  let i = 0
  while (i < src.length) {
    const lt = src.indexOf('<', i)
    if (lt < 0) {
      if (i < src.length) yield { kind: 'text', text: src.slice(i) }
      return
    }
    if (lt > i) yield { kind: 'text', text: src.slice(i, lt) }
    if (src.startsWith('<!--', lt)) {
      const end = src.indexOf('-->', lt + 4)
      i = end < 0 ? src.length : end + 3
      continue
    }
    if (src.startsWith('<![CDATA[', lt)) {
      const end = src.indexOf(']]>', lt + 9)
      yield { kind: 'text', text: src.slice(lt + 9, end < 0 ? src.length : end) }
      i = end < 0 ? src.length : end + 3
      continue
    }
    if (src.startsWith('<?', lt) || src.startsWith('<!', lt)) {
      const end = src.indexOf('>', lt + 2)
      i = end < 0 ? src.length : end + 1
      continue
    }
    let j = lt + 1
    let quote = ''
    for (; j < src.length; j += 1) {
      const ch = src[j]
      if (quote !== '') {
        if (ch === quote) quote = ''
      } else if (ch === '"' || ch === "'") quote = ch
      else if (ch === '>') break
    }
    const raw = src.slice(lt + 1, j)
    const closing = raw.startsWith('/')
    const selfClosing = raw.endsWith('/')
    const body = raw.replace(/^\/+/u, '').replace(/\/+$/u, '').trim()
    const space = body.search(/\s/u)
    const name = space < 0 ? body : body.slice(0, space)
    const attrs = space < 0 ? '' : body.slice(space)
    yield { kind: closing ? 'close' : 'open', name, attrs, selfClosing }
    i = j + 1
  }
}

/** Characters that need escaping when an attribute name becomes a regexp. */
const REGEXP_SPECIAL = /[.*+?^$(){}|[\]\\]/gu

const ATTR_CACHE = new Map()

/** Read one attribute value from a raw attribute string. */
function attr(attrs, name) {
  let re = ATTR_CACHE.get(name)
  if (re === undefined) {
    re = new RegExp('(?:^|\\s)' + name.replace(REGEXP_SPECIAL, '\\$&') + '\\s*=\\s*"([^"]*)"', 'u')
    ATTR_CACHE.set(name, re)
  }
  const m = re.exec(attrs)
  return m === null ? undefined : m[1]
}

/** Decode the five XML entities and numeric character references. */
export function decodeXml(text) {
  return text.replace(/&(#[0-9]+|#x[0-9a-fA-F]+|lt|gt|amp|quot|apos);/gu, (whole, entity) => {
    if (entity === 'lt') return '<'
    if (entity === 'gt') return '>'
    if (entity === 'amp') return '&'
    if (entity === 'quot') return '"'
    if (entity === 'apos') return "'"
    const hex = entity[1] === 'x' || entity[1] === 'X'
    const code = hex ? Number.parseInt(entity.slice(2), 16) : Number.parseInt(entity.slice(1), 10)
    return Number.isFinite(code) && code > 0 && code <= 0x10ffff ? String.fromCodePoint(code) : whole
  })
}

/** Parse a relationship part into id -> relationship. */
function parseRels(xml) {
  const map = new Map()
  if (xml === undefined) return map
  for (const event of xmlEvents(xml)) {
    if (event.kind !== 'open' || event.name !== 'Relationship') continue
    const id = attr(event.attrs, 'Id')
    const target = attr(event.attrs, 'Target')
    if (id === undefined || target === undefined) continue
    map.set(id, {
      target: target.replace(/^\.\//u, ''),
      external: attr(event.attrs, 'TargetMode') === 'External',
    })
  }
  return map
}

/** Resolve a relationship target against the part that owns it. */
function resolvePart(fromPart, target) {
  if (/^[a-zA-Z][a-zA-Z0-9+.-]*:/u.test(target)) return target
  // A leading slash is package-absolute (some producers write '/xl/...').
  if (target.startsWith('/')) return normalizePart(target.slice(1))
  const base = fromPart.includes('/') ? fromPart.slice(0, fromPart.lastIndexOf('/') + 1) : ''
  return normalizePart(base + target)
}

/** Collapse '.' and '..' segments of a package path. */
function normalizePart(path) {
  const out = []
  for (const segment of path.split('/')) {
    if (segment === '' || segment === '.') continue
    if (segment === '..') out.pop()
    else out.push(segment)
  }
  return out.join('/')
}

/** Sibling relationships part of an OOXML part. */
function relsPartOf(part) {
  const slash = part.lastIndexOf('/')
  const dir = slash < 0 ? '' : part.slice(0, slash + 1)
  const file = slash < 0 ? part : part.slice(slash + 1)
  return dir + '_rels/' + file + '.rels'
}

/** Lower-case media type of a part, without the dot. */
function mediaType(part) {
  const dot = part.lastIndexOf('.')
  if (dot < 0) return 'unknown'
  const ext = part.slice(dot + 1).toLowerCase()
  if (ext === 'jpg' || ext === 'jpe') return 'jpeg'
  if (ext === 'tif') return 'tiff'
  return ext
}

/* ------------------------------------------------------------------ *
 * PDF page tree, content streams and per-page images
 * ------------------------------------------------------------------ */

/** Decoded content bytes one PDF scan will spend before it stops reporting. */
const MAX_PDF_STREAM_BYTES = 96 * 1024 * 1024

/** Map a PDF image filter chain to a familiar image type. */
function pdfImageType(filters) {
  if (filters.includes('DCTDecode')) return 'jpeg'
  if (filters.includes('JPXDecode')) return 'jpeg2000'
  if (filters.includes('CCITTFaxDecode')) return 'tiff'
  if (filters.includes('JBIG2Decode')) return 'jbig2'
  if (filters.includes('FlateDecode')) return 'png'
  return 'raw'
}

/** Extract a balanced dictionary literal that starts at 'from'. */
function readDictLiteral(src, from) {
  const start = src.indexOf('<<', from)
  if (start < 0) return undefined
  let depth = 0
  let i = start
  while (i < src.length) {
    if (src.startsWith('<<', i)) { depth += 1; i += 2; continue }
    if (src.startsWith('>>', i)) {
      depth -= 1
      i += 2
      if (depth === 0) return { body: src.slice(start + 2, i - 2), end: i }
      continue
    }
    if (src[i] === '(') {
      let nesting = 1
      i += 1
      while (i < src.length && nesting > 0) {
        if (src[i] === '\\') i += 2
        else {
          if (src[i] === '(') nesting += 1
          else if (src[i] === ')') nesting -= 1
          i += 1
        }
      }
      continue
    }
    i += 1
  }
  return undefined
}

/** Read a named value that is either an inline dictionary or an indirect reference. */
function valueOf(dict, key) {
  const re = new RegExp('/' + key + '\\b\\s*', 'u')
  const m = re.exec(dict)
  if (m === null) return undefined
  const at = m.index + m[0].length
  if (dict.startsWith('<<', at)) {
    const literal = readDictLiteral(dict, at)
    return literal === undefined ? undefined : { kind: 'dict', body: literal.body }
  }
  const ref = /^(\d+)\s+\d+\s+R/u.exec(dict.slice(at))
  return ref === null ? undefined : { kind: 'ref', num: Number(ref[1]) }
}

/** Read a reference array such as '/Kids [ 1 0 R 2 0 R ]'. */
function refsOf(dict, key) {
  const re = new RegExp('/(?:' + key + ')\\s*\\[([\\s\\S]*?)\\]', 'u')
  const m = re.exec(dict)
  if (m === null) return undefined
  const out = []
  for (const hit of m[1].matchAll(/(\d+)\s+\d+\s+R/gu)) out.push(Number(hit[1]))
  return out
}

/** Read a number array such as a MediaBox. */
function numbersOf(dict, key) {
  const re = new RegExp('/(?:' + key + ')\\s*\\[([^\\]]*)\\]', 'u')
  const m = re.exec(dict)
  if (m === null) return undefined
  return m[1].trim().split(/\s+/u).map(Number).filter((n) => Number.isFinite(n))
}

/** Read one indirect reference value ('/Key 12 0 R') as a number. */
function refNumber(dict, key) {
  const m = new RegExp('/' + key + '\\s+(\\d+)\\s+\\d+\\s+R', 'u').exec(dict)
  return m === null ? undefined : Number(m[1])
}

/**
 * Index every plain object of a PDF. Stream payloads stay as offsets: the
 * offset scan means binary image data is never parsed as an object.
 * @param {string} src - the file decoded as latin1, one char per byte.
 */
function scanPdfObjects(src) {
  const objects = new Map()
  const pattern = /(\d+)\s+(\d+)\s+obj\b/gu
  let match
  while ((match = pattern.exec(src)) !== null) {
    const bodyStart = match.index + match[0].length
    const streamAt = src.indexOf('stream', bodyStart)
    const endAt = src.indexOf('endobj', bodyStart)
    const hasStream = streamAt >= 0 && (endAt < 0 || streamAt < endAt)
    const dictStop = hasStream ? streamAt : endAt < 0 ? src.length : endAt
    let dataStart
    let endStreamAt
    if (hasStream) {
      dataStart = streamAt + 6
      if (src[dataStart] === '\r') dataStart += 1
      if (src[dataStart] === '\n') dataStart += 1
      endStreamAt = src.indexOf('endstream', dataStart)
      const after = endStreamAt < 0 ? -1 : src.indexOf('endobj', endStreamAt)
      pattern.lastIndex = after < 0 ? src.length : after + 6
    }
    objects.set(Number(match[1]), { dict: src.slice(bodyStart, dictStop), dataStart, endStreamAt })
  }
  return objects
}

/**
 * Parse a PDF's page tree and each page's drawn images.
 *
 * A picture counts for a page only when that page's content stream actually
 * invokes its XObject with 'Do'. Reading '/Resources' instead would count a
 * picture on every page, because LibreOffice shares one resource dictionary
 * across pages.
 * @param {Buffer} buf - complete PDF bytes.
 */
export function parsePdf(buf) {
  const src = buf.toString('latin1')
  const objects = scanPdfObjects(src)
  const budget = { left: MAX_PDF_STREAM_BYTES }

  const readStream = (object) => {
    if (object === undefined || object.dataStart === undefined) return undefined
    let length
    // Test the indirect form first: '/Length 34 0 R' would otherwise be read
    // as the literal length 3 and truncate every stream it appears in.
    const indirect = /\/Length\s+(\d+)\s+\d+\s+R/u.exec(object.dict)
    if (indirect !== null) {
      const target = objects.get(Number(indirect[1]))
      const value = target === undefined ? null : /^\s*(\d+)\s*$/u.exec(target.dict)
      if (value !== null) length = Number(value[1])
    } else {
      const direct = /\/Length\s+(\d+)/u.exec(object.dict)
      if (direct !== null) length = Number(direct[1])
    }
    const end = object.endStreamAt < 0 ? object.dataStart : object.endStreamAt
    const stop = length === undefined ? end : Math.min(object.dataStart + length, end)
    if (stop <= object.dataStart) return ''
    const raw = Buffer.from(src.slice(object.dataStart, stop), 'latin1')
    budget.left -= raw.length
    if (!/\/Filter\b/u.test(object.dict)) return raw.toString('latin1')
    if (!/\/FlateDecode\b/u.test(object.dict)) return undefined
    try {
      const inflated = inflateRawSync(raw)
      budget.left -= inflated.length
      return inflated.toString('latin1')
    } catch {
      try {
        const inflated = inflateSync(raw)
        budget.left -= inflated.length
        return inflated.toString('latin1')
      } catch { return undefined }
    }
  }

  const resolveResources = (dict, inherited) => {
    const value = valueOf(dict, 'Resources')
    if (value === undefined) return inherited
    if (value.kind === 'dict') return value.body
    const target = objects.get(value.num)
    return target === undefined ? inherited : target.dict
  }

  const imageOf = (dict) => {
    const filters = [...dict.matchAll(/\/(DCTDecode|JPXDecode|CCITTFaxDecode|JBIG2Decode|FlateDecode|LZWDecode|RunLengthDecode)\b/gu)].map((hit) => hit[1])
    const width = Number(/\/(?:Width|W)\s+(\d+)/u.exec(dict)?.[1] ?? 0)
    const height = Number(/\/(?:Height|H)\s+(\d+)/u.exec(dict)?.[1] ?? 0)
    return { type: pdfImageType(filters), width: width || undefined, height: height || undefined }
  }

  /** Images a content stream actually paints, following Form XObjects. */
  const paintedImages = (resources, content, depth) => {
    const found = []
    if (content === undefined || content === '' || depth > 3) return found
    const value = valueOf(resources ?? '', 'XObject')
    if (value === undefined) return found
    const body = value.kind === 'dict' ? value.body : objects.get(value.num)?.dict
    if (body === undefined) return found
    const map = new Map()
    for (const hit of body.matchAll(/\/([^\s/<>[\]()]+)\s+(\d+)\s+\d+\s+R/gu)) map.set(hit[1], Number(hit[2]))
    for (const hit of content.matchAll(/\/([^\s/<>[\](){}]+)\s+Do\b/gu)) {
      const target = objects.get(map.get(hit[1]) ?? -1)
      if (target === undefined) continue
      if (/\/Subtype\s*\/Image\b/u.test(target.dict)) found.push(imageOf(target.dict))
      else if (/\/Subtype\s*\/Form\b/u.test(target.dict)) {
        const inner = readStream(target)
        found.push(...paintedImages(resolveResources(target.dict, resources), inner, depth + 1))
      }
    }
    return found
  }

  const contentOf = (dict) => {
    const parts = []
    const array = /\/Contents\s*\[([^\]]*)\]/u.exec(dict)
    if (array !== null) {
      for (const hit of array[1].matchAll(/(\d+)\s+\d+\s+R/gu)) parts.push(Number(hit[1]))
    } else {
      const single = refNumber(dict, 'Contents')
      if (single !== undefined) parts.push(single)
    }
    let text = ''
    for (const part of parts) text += (readStream(objects.get(part)) ?? '') + '\n'
    return text
  }

  const pages = []
  const seen = new Set()
  const walk = (nodeNumber, inherited) => {
    if (nodeNumber === undefined || seen.has(nodeNumber)) return
    const node = objects.get(nodeNumber)
    if (node === undefined) return
    seen.add(nodeNumber)
    const mediaBox = numbersOf(node.dict, 'MediaBox') ?? inherited.mediaBox
    const resources = resolveResources(node.dict, inherited.resources)
    if (/\/Type\s*\/Pages\b/u.test(node.dict)) {
      for (const kid of refsOf(node.dict, 'Kids') ?? []) walk(kid, { mediaBox, resources })
      return
    }
    if (!/\/Type\s*\/Page\b/u.test(node.dict)) return
    pages.push({ mediaBox, resources, dict: node.dict })
  }

  const rootMatch = /\/Root\s+(\d+)\s+\d+\s+R/u.exec(src)
  if (rootMatch !== null) {
    const root = objects.get(Number(rootMatch[1]))
    if (root !== undefined) walk(refNumber(root.dict, 'Pages'), { mediaBox: undefined, resources: undefined })
  }
  if (pages.length === 0) {
    for (const [, node] of objects) {
      if (!/\/Type\s*\/Page\b/u.test(node.dict)) continue
      pages.push({ mediaBox: numbersOf(node.dict, 'MediaBox'), resources: resolveResources(node.dict, undefined), dict: node.dict })
    }
  }

  let incomplete = false
  const result = pages.map((page, index) => {
    const images = paintedImages(page.resources, contentOf(page.dict), 0)
    if (budget.left <= 0) incomplete = true
    const box = page.mediaBox
    return {
      n: index + 1,
      imageCount: images.length,
      types: [...new Set(images.map((image) => image.type))],
      size: box === undefined || box.length < 4 ? '' : mmOf(Math.abs(box[2] - box[0])) + '×' + mmOf(Math.abs(box[3] - box[1])) + 'mm',
    }
  })
  return { pageCount: result.length, pages: result, incomplete }
}

function mmOf(points) {
  return Math.round((points / 72) * 25.4)
}
/* ------------------------------------------------------------------ *
 * DOCX
 * ------------------------------------------------------------------ */

/** Word heading style names that count toward the document outline. */
const HEADING_STYLE = /^(?:Heading|heading|标题)\s*(\d)$/u

function parseDocx(zip) {
  const documentPart = 'word/document.xml'
  const xml = zip.text(documentPart)
  if (xml === undefined) throw new Error('the package has no word/document.xml')
  const rels = parseRels(zip.text('word/_rels/document.xml.rels'))
  const app = zip.text('docProps/app.xml') ?? ''
  const model = { sections: [], tables: [], images: [], paragraphs: 0, chars: 0, headings: [], reportedPages: undefined }
  const reported = /<Pages>(\d+)<\/Pages>/u.exec(app)
  if (reported !== null) model.reportedPages = Number(reported[1])

  const tableStack = []
  let section
  let paragraph
  let inText = false
  let textBuffer = ''
  let drawing
  let placement

  const pushImage = (rid, where, extent) => {
    const rel = rid === undefined ? undefined : rels.get(rid)
    const part = rel === undefined || rel.external ? undefined : resolvePart(documentPart, rel.target)
    model.images.push({
      type: part === undefined ? 'unknown' : mediaType(part),
      placement: where,
      width: extent?.cx,
      height: extent?.cy,
    })
  }

  for (const event of xmlEvents(xml)) {
    if (event.kind === 'text') {
      if (inText) textBuffer += event.text
      continue
    }
    const { name, attrs, selfClosing } = event
    const opening = event.kind === 'open'

    if (name === 'w:tbl') {
      if (opening && !selfClosing) tableStack.push({ rows: 0, maxCells: 0, rowCells: 0 })
      else if (!opening) {
        const frame = tableStack.pop()
        if (frame !== undefined && tableStack.length === 0) model.tables.push({ rows: frame.rows, cols: frame.maxCells })
      }
      continue
    }
    if (name === 'w:tr' && opening && !selfClosing) {
      const frame = tableStack[tableStack.length - 1]
      if (frame !== undefined) { frame.rows += 1; frame.rowCells = 0 }
      continue
    }
    if (name === 'w:tc' && opening) {
      const frame = tableStack[tableStack.length - 1]
      if (frame !== undefined) { frame.rowCells += 1; frame.maxCells = Math.max(frame.maxCells, frame.rowCells) }
      continue
    }
    if (name === 'w:sectPr') {
      if (opening) {
        section = { breakType: 'nextPage', orientation: 'portrait', width: 0, height: 0 }
        if (selfClosing) { model.sections.push(section); section = undefined }
      } else if (section !== undefined) { model.sections.push(section); section = undefined }
      continue
    }
    if (section !== undefined && opening) {
      if (name === 'w:type') section.breakType = attr(attrs, 'w:val') ?? section.breakType
      if (name === 'w:pgSz') {
        section.width = Number(attr(attrs, 'w:w') ?? 0)
        section.height = Number(attr(attrs, 'w:h') ?? 0)
        section.orientation = attr(attrs, 'w:orient') === 'landscape' ? 'landscape' : 'portrait'
      }
    }
    if (name === 'w:p') {
      if (opening) {
        if (tableStack.length === 0) model.paragraphs += 1
        paragraph = { style: '', outlineLevel: undefined, text: '' }
      } else if (paragraph !== undefined) {
        const trimmed = paragraph.text.trim()
        const level = Number(HEADING_STYLE.exec(paragraph.style)?.[1] ?? NaN)
        if ((Number.isFinite(level) || (paragraph.outlineLevel !== undefined && paragraph.outlineLevel < 9)) && trimmed !== '') {
          model.headings.push(trimmed.slice(0, 80))
        }
        paragraph = undefined
      }
      continue
    }
    if (paragraph !== undefined && opening) {
      if (name === 'w:pStyle') paragraph.style = attr(attrs, 'w:val') ?? ''
      if (name === 'w:outlineLvl') paragraph.outlineLevel = Number(attr(attrs, 'w:val') ?? 9)
      if (name === 'w:drawing' || name === 'w:pict') { drawing = { cx: undefined, cy: undefined }; placement = 'inline' }
      if (name === 'wp:inline') placement = 'inline'
      if (name === 'wp:anchor') placement = 'floating'
      if (name === 'wp:extent' && drawing !== undefined) {
        drawing.cx = Number(attr(attrs, 'cx') ?? 0) || undefined
        drawing.cy = Number(attr(attrs, 'cy') ?? 0) || undefined
      }
      if ((name === 'a:blip' || name === 'v:imagedata') && drawing !== undefined) {
        pushImage(name === 'a:blip' ? attr(attrs, 'r:embed') : attr(attrs, 'r:id'), placement === 'floating' ? 'floating' : 'inline', drawing)
      }
    }
    if ((name === 'w:drawing' || name === 'w:pict') && !opening) { drawing = undefined; placement = undefined }
    if (name === 'w:t') {
      if (opening && !selfClosing) { inText = true; textBuffer = '' }
      else if (!opening) {
        inText = false
        const decoded = decodeXml(textBuffer)
        if (paragraph !== undefined) { paragraph.text += decoded; model.chars += decoded.length }
        textBuffer = ''
      }
    }
    if ((name === 'w:tab' || name === 'w:br') && opening && paragraph !== undefined) paragraph.text += ' '
  }

  if (section !== undefined) model.sections.push(section)
  if (model.sections.length === 0) model.sections.push({ breakType: 'nextPage', orientation: 'portrait', width: 0, height: 0 })
  return model
}

/* ------------------------------------------------------------------ *
 * PPTX
 * ------------------------------------------------------------------ */

function parsePptx(zip) {
  const presentation = zip.text('ppt/presentation.xml')
  if (presentation === undefined) throw new Error('the package has no ppt/presentation.xml')
  const rels = parseRels(zip.text('ppt/_rels/presentation.xml.rels'))
  const order = []
  for (const hit of presentation.matchAll(/<p:sldId\b[^>]*\br:id="([^"]+)"/gu)) order.push(hit[1])
  const sizeMatch = /<p:sldSz\b([^>]*?)\/?>/u.exec(presentation)
  const width = Number(sizeMatch === null ? 0 : attr(sizeMatch[1], 'cx') ?? 0)

  const slides = []
  for (const [index, rid] of order.entries()) {
    const rel = rels.get(rid)
    if (rel === undefined || rel.external) continue
    const part = resolvePart('ppt/presentation.xml', rel.target)
    const xml = zip.text(part)
    if (xml === undefined) continue
    const slideRels = parseRels(zip.text(relsPartOf(part)))
    const types = []
    for (const hit of xml.matchAll(/<a:blip\b[^>]*\br:embed="([^"]+)"/gu)) {
      const media = slideRels.get(hit[1])
      types.push(media === undefined ? 'unknown' : mediaType(resolvePart(part, media.target)))
    }
    let chars = 0
    for (const hit of xml.matchAll(/<a:t\b[^>]*>([\s\S]*?)<\/a:t>/gu)) chars += decodeXml(hit[1]).length
    slides.push({
      n: index + 1,
      imageCount: types.length,
      types: [...new Set(types)],
      tables: [...xml.matchAll(/<a:tbl>/gu)].length,
      charts: [...xml.matchAll(/<c:chart\b/gu)].length,
      chars,
    })
  }
  return { slides, width: width === 0 ? 0 : Math.round(width / 36000) }
}

/* ------------------------------------------------------------------ *
 * XLSX
 * ------------------------------------------------------------------ */

function parseXlsx(zip) {
  const workbook = zip.text('xl/workbook.xml')
  if (workbook === undefined) throw new Error('the package has no xl/workbook.xml')
  const rels = parseRels(zip.text('xl/_rels/workbook.xml.rels'))
  const sheets = []
  let index = 0
  for (const hit of workbook.matchAll(/<sheet\b([^>]*?)\/?>/gu)) {
    index += 1
    const attrs = hit[1]
    const name = attr(attrs, 'name') ?? 'Sheet' + index
    const hidden = (attr(attrs, 'state') ?? 'visible') !== 'visible'
    const rid = attr(attrs, 'r:id')
    const rel = rid === undefined ? undefined : rels.get(rid)
    const part = rel === undefined ? undefined : resolvePart('xl/workbook.xml', rel.target)
    const xml = part === undefined ? undefined : zip.text(part)
    if (xml === undefined || part === undefined) {
      sheets.push({ name, hidden, rows: 0, cols: 0, images: 0, charts: 0, tables: 0, merges: 0 })
      continue
    }
    const dimension = /<dimension\b[^>]*\bref="([^"]+)"/u.exec(xml)
    const sheetRels = parseRels(zip.text(relsPartOf(part)))
    let images = 0
    let charts = 0
    for (const drawing of xml.matchAll(/<drawing\b[^>]*\br:id="([^"]+)"/gu)) {
      const drawingRel = sheetRels.get(drawing[1])
      if (drawingRel === undefined) continue
      const drawingPart = resolvePart(part, drawingRel.target)
      const drawingXml = zip.text(drawingPart) ?? ''
      images += [...drawingXml.matchAll(/<(?:xdr:)?pic\b/gu)].length
      charts += [...drawingXml.matchAll(/<c:chart\b/gu)].length
    }
    const size = dimension === null ? { rows: 0, cols: 0 } : dimensionSize(dimension[1])
    sheets.push({
      name,
      hidden,
      dimension: dimension === null ? '' : dimension[1],
      rows: size.rows,
      cols: size.cols,
      images,
      charts,
      tables: [...xml.matchAll(/<tablePart\b/gu)].length,
      merges: [...xml.matchAll(/<mergeCell\b/gu)].length,
    })
  }
  return { sheets }
}

function dimensionSize(ref) {
  if (!ref.includes(':')) {
    const single = /^([A-Z]+)(\d+)$/u.exec(ref)
    return single === null ? { rows: 0, cols: 0 } : { rows: 1, cols: columnIndex(single[1]) }
  }
  const pair = ref.split(':')
  const a = /^([A-Z]+)(\d+)$/u.exec(pair[0])
  const b = /^([A-Z]+)(\d+)$/u.exec(pair[1])
  if (a === null || b === null || pair.length !== 2) return { rows: 0, cols: 0 }
  return { rows: Math.abs(Number(b[2]) - Number(a[2])) + 1, cols: Math.abs(columnIndex(b[1]) - columnIndex(a[1])) + 1 }
}

function columnIndex(letters) {
  let value = 0
  for (const ch of letters) value = value * 26 + (ch.charCodeAt(0) - 64)
  return value
}

/* ------------------------------------------------------------------ *
 * Public scan entry
 * ------------------------------------------------------------------ */

/** Document families the outline supports, by lower-case extension. */
export const DOCUMENT_FAMILIES = {
  doc: 'word', docx: 'word', odt: 'word',
  xls: 'excel', xlsx: 'excel', ods: 'excel',
  ppt: 'powerpoint', pptx: 'powerpoint', odp: 'powerpoint',
  pdf: 'pdf',
}

/** Lower-case extension of a path. */
export function extensionOf(file) {
  const dot = file.lastIndexOf('.')
  return dot < 0 ? '' : file.slice(dot + 1).toLowerCase()
}

/**
 * Scan one document's structure without rendering it.
 *
 * 'pages' carries per-page picture counts for Word documents and PDFs, the two
 * families whose page numbers the renderer addresses.
 * @param {Buffer} bytes - complete document bytes.
 * @param {string} ext - lower-case extension.
 * @param {{ pageCount?: number, pdf?: { pageCount: number, pages: any[] } }} [layout] - pagination facts already obtained from the Office service.
 */
export function scanStructure(bytes, ext, layout) {
  const family = DOCUMENT_FAMILIES[ext]
  if (family === undefined) throw new Error('unsupported document extension "' + ext + '"')
  const model = {
    format: ext,
    family,
    pageCount: 0,
    pageCountSource: '',
    pages: [],
    slides: [],
    sheets: [],
    sections: [],
    tables: [],
    images: [],
    headings: [],
    paragraphs: 0,
    chars: 0,
    reportedPages: undefined,
    notes: [],
  }

  if (family === 'pdf') {
    const pdf = parsePdf(bytes)
    model.pageCount = pdf.pageCount
    model.pageCountSource = 'pdf'
    model.pages = pdf.pages
    if (pdf.incomplete) model.notes.push('页级扫描未完成（内容流过大或无法解码）：图片统计可能偏少，按页渲染前请先确认。')
    return model
  }

  const zip = openZip(bytes)
  if (family === 'word' && ext === 'docx') {
    const word = parseDocx(zip)
    model.sections = word.sections
    model.tables = word.tables
    model.images = word.images
    model.headings = word.headings
    model.paragraphs = word.paragraphs
    model.chars = word.chars
    model.reportedPages = word.reportedPages
  } else if (family === 'powerpoint' && ext === 'pptx') {
    const deck = parsePptx(zip)
    model.slides = deck.slides
    model.pageCount = deck.slides.length
    model.pageCountSource = 'ooxml'
    if (deck.width > 0) model.size = deck.width + 'mm 宽'
  } else if (family === 'excel' && ext === 'xlsx') {
    const book = parseXlsx(zip)
    model.sheets = book.sheets
    model.pageCount = book.sheets.filter((sheet) => !sheet.hidden).length
    model.pageCountSource = 'ooxml'
  } else {
    model.notes.push('旧版或开放文档格式（' + ext + '）无法结构化解析：只能给转换引擎报告的页数，要看版面请整页渲染。')
  }

  if (layout !== undefined) {
    if (layout.pdf !== undefined && layout.pdf.pageCount > 0) {
      model.pages = layout.pdf.pages
      if (layout.pdf.incomplete) model.notes.push('页级扫描未完成（内容流过大或无法解码）：图片统计可能偏少，按页渲染前请先确认。')
      if (model.pageCount === 0) {
        model.pageCount = layout.pdf.pageCount
        model.pageCountSource = 'converter'
      } else if (layout.pdf.pageCount !== model.pageCount) {
        model.notes.push('转换引擎分页为 ' + layout.pdf.pageCount + ' 页，结构解析得到 ' + model.pageCount + ' 页，以结构解析为准。')
      }
    }
    if (model.pageCount === 0 && layout.pageCount !== undefined && layout.pageCount > 0) {
      model.pageCount = layout.pageCount
      model.pageCountSource = 'converter'
    }
  }

  const totalOnPages = model.pages.reduce((sum, page) => sum + page.imageCount, 0)
  if (model.pages.length > 0 && model.images.length > 0 && model.images.length === totalOnPages) {
    let cursor = 0
    for (const page of model.pages) {
      if (page.imageCount > 0) {
        const slice = model.images.slice(cursor, cursor + page.imageCount).map((image) => image.type).filter((type) => type !== 'unknown')
        if (slice.length > 0) page.types = [...new Set(slice)]
      }
      cursor += page.imageCount
    }
  }
  return model
}

/**
 * Normalize converted Word text so line-oriented readers see paragraphs.
 *
 * LibreOffice writes the text filter with a bare CR between paragraphs; the file
 * contains no LF at all, so grep and every other line reader see one enormous
 * line. Replacing CR/CRLF with LF keeps the paragraph boundaries the document
 * has, and the returned statistics let a caller detect a still-broken channel.
 * @param {string} text - converted text.
 */
export function normalizeText(text) {
  const body = text.replace(/^\uFEFF/u, '').replace(/\r\n?/gu, '\n')
  const lines = body.split('\n')
  while (lines.length > 0 && lines[lines.length - 1] === '') lines.pop()
  let chars = 0
  let maxLineLength = 0
  for (const line of lines) { chars += line.length; if (line.length > maxLineLength) maxLineLength = line.length }
  const collapsed = (lines.length <= 1 && chars > 200) || maxLineLength > 2000
  let reason = ''
  if (collapsed) {
    reason = lines.length <= 1 ? '整篇只有 1 行且超过 200 字，段落边界丢失' : '存在超过 2000 字的单行，按行读取会被截断'
  }
  return { text: lines.join('\n'), lines: lines.length, chars, maxLineLength, textChannel: collapsed ? 'suspect' : 'ok', reason }
}
