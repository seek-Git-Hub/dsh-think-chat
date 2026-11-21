/**
 * Read-only Office views for the 仅对话 presets: a whole-document structure
 * manifest and a paragraph-preserving text channel.
 *
 * Why these exist:
 * - office_convert writes Word text with bare CR between paragraphs, so a line
 *   reader (grep included) sees the whole document as one enormous line. That is
 *   an official converter detail this bundle must not change, so the text
 *   channel normalizes CR/CRLF to LF itself and reports line statistics that
 *   make a still-broken channel detectable instead of silently guessed.
 * - Deciding "does this document have pictures" from text is impossible, and
 *   rendering every page to find out spends vision tokens on text-only pages.
 *   office_outline answers it from the container and the page tree, without
 *   producing a single image.
 *
 * Both tools are read-only: they never write inside the workspace, never run a
 * shell, never reach the network, and never turn the preset into a document
 * editor. They reuse the official Office service for pagination and the text
 * conversion, so the file formats supported here are exactly the ones that
 * service already supports.
 *
 * @module dsh-better-chat/office-read
 */

import { readFile, writeFile } from 'node:fs/promises'
import { basename, dirname, join } from 'node:path'
import { defineTool } from '@deepseek-ai/dsh-tools'
import { DOCUMENT_FAMILIES, extensionOf, normalizeText, parsePdf, scanStructure } from './office-structure.js'

/** Largest source document a scan or text conversion will read. */
const MAX_INPUT_BYTES = 64 * 1024 * 1024
/** Largest converted text result this channel will load. */
const MAX_TEXT_BYTES = 8 * 1024 * 1024
/** Characters one office_text slice returns before it asks for the next offset. */
const MAX_SLICE_CHARS = 12000
/** Word-family extensions that have a text channel at all. */
const WORD_FAMILY = new Set(['doc', 'docx', 'odt'])

const KIND_LABEL = {
  word: 'Word 文档',
  excel: 'Excel 工作簿',
  powerpoint: 'PowerPoint 演示文稿',
  pdf: 'PDF 文档',
}

const PAGE_SOURCE_LABEL = {
  converter: '转换引擎分页',
  pdf: 'PDF 页树',
  ooxml: '文件结构',
}

/** Read a mounted service or fail with an actionable message. */
function service(ctx, name, consequence) {
  const value = ctx.get(name)
  if (value === undefined) throw new Error(consequence)
  return value
}

/**
 * Resolve the model's input path through the ordinary filesystem authorization.
 * @returns the fs service, calling agent, target, stat and resolved directory.
 */
async function resolveInput(ctx, exec, input) {
  const fsService = service(ctx, 'fs', '无法读取 "' + input + '"：文件系统服务不可用')
  const agent = exec.agent
  if (agent === undefined) throw new Error('无法读取 "' + input + '"：该工具只在一次 Agent 步骤中执行')
  const directory = ctx.get('workingDirectory')
  const cwd = directory === undefined ? undefined : await directory.ensure(agent, exec.signal)
  const target = await fsService.resolve(input, cwd === undefined ? { signal: exec.signal } : { cwd, signal: exec.signal })
  const info = await fsService.stat(target, exec.signal)
  if (info === undefined) throw new Error('无法读取 "' + input + '"：文件不存在')
  if (info.type !== 'file') throw new Error('无法读取 "' + input + '"：不是普通文件')
  return { fsService, agent, target, info, cwd }
}

/** Authorized deferred source read for the Office service. */
function sourceOf(fsService, agent, target, info) {
  return {
    key: JSON.stringify([agent.session.id, fsService.processPath(target)]),
    version: info.version,
    ...(info.size === undefined ? {} : { bytes: info.size }),
    read(signal, maxBytes) {
      return fsService.readBytes(target, signal, maxBytes).then((bytes) => ({ bytes, version: info.version }))
    },
  }
}

/** Mounted Office conversion service, or an actionable failure. */
function officeService(ctx) {
  return service(ctx, 'officeKit', '文档转换服务不可用：officeKit 未挂载，无法读取 Office 文档')
}

/** Short human text for a thrown value. */
function messageOf(error) {
  return error instanceof Error ? error.message : String(error)
}

/* ------------------------------------------------------------------ *
 * Manifest rendering
 * ------------------------------------------------------------------ */

/** Twips (Word section sizes) to millimetres. */
function twipsToMm(twips) {
  return Math.round((twips / 1440) * 25.4)
}

/** Join a bounded, de-duplicated list for display. */
function joinList(items, limit) {
  const unique = [...new Set(items.filter((item) => item !== ''))]
  const head = unique.slice(0, limit)
  return unique.length > limit ? head.join(' / ') + ' / …（共 ' + unique.length + ' 项）' : head.join(' / ')
}

/** Word section break kinds in the words a reader expects. */
const BREAK_LABEL = { nextPage: '分页', continuous: '连续', evenPage: '偶数页', oddPage: '奇数页' }

/** An office_render pages argument a caller can paste back verbatim. */
function renderCall(pages) {
  if (pages.length === 0) return '（无）'
  if (pages.length <= 20) return 'pages=[' + pages.join(',') + ']'
  return 'pages=[' + pages.slice(0, 20).join(',') + ',…]（共 ' + pages.length + ' 页，请分批）'
}

/** Every page number of a document, for a render-everything suggestion. */
function allPages(count) {
  const pages = []
  for (let n = 1; n <= count; n += 1) pages.push(n)
  return pages
}

/**
 * Compress consecutive pages that share a picture summary into ranges.
 * @param {Array<{ n: number, imageCount: number, types: string[] }>} pages
 */
function pageRanges(pages) {
  const out = []
  for (const page of pages) {
    const label = page.imageCount === 0 ? '0张' : page.imageCount + '张(' + (page.types.length === 0 ? '类型未知' : page.types.join('/')) + ')'
    const last = out[out.length - 1]
    if (last !== undefined && last.label === label && last.to === page.n - 1) last.to = page.n
    else out.push({ from: page.n, to: page.n, label })
  }
  return out.map((range) => (range.from === range.to ? range.from + '页 ' + range.label : range.from + '-' + range.to + '页 ' + range.label))
}

/** Pictures worth rendering alone, and pages that are text-only. */
function pageDecision(model) {
  if (model.family !== 'word' && model.family !== 'pdf') return { imagePages: [], textPages: [] }
  const imagePages = model.pages.filter((page) => page.imageCount > 0).map((page) => page.n)
  const textPages = model.pages.filter((page) => page.imageCount === 0).map((page) => page.n)
  return { imagePages, textPages }
}

/** The model-facing manifest: facts first, then one executable next step. */
function manifestOf(display, model, decision, advice) {
  const lines = []
  const source = PAGE_SOURCE_LABEL[model.pageCountSource] ?? model.pageCountSource
  lines.push('office_outline · ' + basename(display) + ' · ' + (KIND_LABEL[model.family] ?? model.format))
  lines.push('文件：' + display)
  lines.push('页数：' + (model.pageCount > 0 ? model.pageCount + '（' + source + '）' : '未能确定（转换引擎分页失败）'))

  if (model.family === 'word') {
    const sections = model.sections.map((section, index) => {
      const number = '第' + (index + 1) + '节'
      const size = section.width > 0 ? ' ' + twipsToMm(section.width) + '×' + twipsToMm(section.height) + 'mm' : ''
      return number + ' ' + (BREAK_LABEL[section.breakType] ?? section.breakType) + '·' + (section.orientation === 'landscape' ? '横向' : '纵向') + size
    })
    if (sections.length > 0) lines.push('分节：' + sections.length + ' 节（' + joinList(sections, 6) + '）')
    lines.push('表格：' + (model.tables.length === 0 ? '无' : model.tables.length + ' 个（' + model.tables.slice(0, 6).map((table) => table.rows + '行×' + table.cols + '列').join('；') + '）'))
    const inline = model.images.filter((image) => image.placement !== 'floating').length
    const floating = model.images.length - inline
    const types = [...new Set(model.images.map((image) => image.type))].join('/')
    lines.push('图片：' + (model.images.length === 0 ? '无' : model.images.length + ' 张（' + types + '；内嵌 ' + inline + ' / 浮动 ' + floating + '）'))
    if (model.headings.length > 0) lines.push('标题：' + joinList(model.headings, 8))
    lines.push('正文：' + model.paragraphs + ' 段 · 约 ' + model.chars + ' 字')
    if (model.pages.length > 0) lines.push('每页图片：' + pageRanges(model.pages).join(' · '))
  } else if (model.family === 'pdf') {
    if (model.pages.length > 0) lines.push('每页图片：' + pageRanges(model.pages).join(' · '))
  } else if (model.family === 'powerpoint') {
    if (model.slides.length > 0) {
      lines.push('幻灯片：' + model.slides.length + ' 张（' + model.slides.slice(0, 8).map((slide) => slide.n + ': ' + slide.imageCount + '图/' + slide.tables + '表/' + slide.chars + '字').join('；') + '）')
    }
  } else if (model.family === 'excel') {
    if (model.sheets.length > 0) {
      lines.push('工作表：' + model.sheets.map((sheet) => sheet.name + (sheet.hidden ? '（隐藏）' : '') + ' ' + sheet.rows + '×' + sheet.cols + (sheet.dimension ? ' [' + sheet.dimension + ']' : '') + (sheet.images > 0 ? ' 图' + sheet.images : '') + (sheet.tables > 0 ? ' 表' + sheet.tables : '') + (sheet.merges > 0 ? ' 合并' + sheet.merges : '')).join('；'))
    }
  }

  for (const note of model.notes) lines.push('注意：' + note)
  lines.push('建议：' + advice)
  return lines.join('\n')
}

/** One executable next step derived from the manifest. */
function adviceOf(model, decision) {
  if (model.family === 'pdf') {
    if (model.pageCount === 0) return 'PDF 页数未能解析：直接整页渲染，不要用文本猜测版面。'
    if (decision.imagePages.length === model.pageCount) {
      return 'PDF 没有文本通道：按问题挑页渲染 ' + renderCall(allPages(model.pageCount)) + '，不要全篇渲染后再筛。'
    }
    return 'PDF 没有文本通道（文字也只能靠渲染读）：有图页 ' + renderCall(decision.imagePages) + ' 优先渲染，其余页按需渲染。'
  }
  if (model.family === 'powerpoint') {
    const pages = allPages(model.pageCount)
    return '幻灯片没有文本通道：分批用 office_render 渲染 ' + renderCall(pages) + '，每批不超过 6 张，按问题挑页。'
  }
  if (model.family === 'excel') {
    const names = model.sheets.filter((sheet) => !sheet.hidden).map((sheet) => sheet.name).slice(0, 8).join('、')
    return '工作表没有文本通道：用 office_render({ input, sheet: "表名" }) 逐表渲染需要的区域（可见表：' + (names === '' ? '无' : names) + '）；隐藏表默认不渲染。'
  }
  if (model.pageCount === 0) return '页数未能确定：先用 office_text 取文本并看 textChannel，再决定按页渲染；不要猜页数。'
  if (decision.imagePages.length === 0) {
    const tables = model.tables.length > 0 ? '表格如需看对齐/合并效果，再单独渲染对应页。' : ''
    return '全篇无图片：直接用 office_text 提取文本分段读，不要渲染。' + tables
  }
  if (decision.imagePages.length === model.pageCount) {
    return '每页都有图：只渲染有图页 ' + renderCall(decision.imagePages) + '；文字用 office_text 读，不要用渲染结果推断文本。'
  }
  return '有图页 ' + renderCall(decision.imagePages) + ' → 只用 office_render 渲染这些页；其余纯文本页 → 用 office_text 提取文本；两条路按页选择，不得混用。'
}

/* ------------------------------------------------------------------ *
 * Tool registration
 * ------------------------------------------------------------------ */

/**
 * Register office_outline and office_text in the calling scope.
 * @param ctx - preset scope; tools are registered here and disappear with it.
 */
export function registerOfficeReadTools(ctx) {
  ctx.tools.register(defineTool({
    name: 'office_outline',
    description: [
      '读取 .docx/.xlsx/.pptx/.pdf 的整篇结构清单，不生成图片：页数、每页图片数量与类型、表格、分节、标题、文字量。',
      '读任何 Office/PDF 文档前必须先调用本工具取得清单；没有清单不得开始文本检索或渲染。',
      'Word/PDF 的页号与 office_render 的 pages 完全一致（同一转换引擎分页）。',
    ].join('\n'),
    parameters: {
      input: { type: 'string', required: true, description: '文档路径；相对路径按当前工作目录解析。' },
    },
    output: {
      schema: {
        type: 'object',
        additionalProperties: false,
        properties: {
          path: { type: 'string', required: true },
          format: { type: 'string', required: true },
          family: { type: 'string', required: true },
          pageCount: { type: 'integer', required: true },
          pageCountSource: { type: 'string', required: true },
          imagePages: { type: 'array', required: true, items: { type: 'integer' } },
          textPages: { type: 'array', required: true, items: { type: 'integer' } },
          advice: { type: 'string', required: true },
          manifest: { type: 'string', required: true },
          notes: { type: 'array', required: true, items: { type: 'string' } },
        },
      },
      render: (_args, value) => [{ type: 'text', text: value.manifest }],
    },
    async execute(args, exec) {
      const ext = extensionOf(args.input)
      const family = DOCUMENT_FAMILIES[ext]
      if (family === undefined) {
        throw new Error('无法解析 "' + args.input + '"：只支持 doc/docx/odt/xls/xlsx/ods/ppt/pptx/odp/pdf，不支持 ".' + ext + '"')
      }
      const { fsService, agent, target, info } = await resolveInput(ctx, exec, args.input)
      const bytes = Buffer.from(await fsService.readBytes(target, exec.signal, MAX_INPUT_BYTES))
      let layout
      if (family === 'word') {
        // Word pagination and per-page pictures only exist after layout, so ask
        // the official converter for a PDF and read its page tree. No image is
        // produced, and the PDF is never shown to the model.
        try {
          const kit = officeService(ctx)
          const converted = await kit.convertTo({
            extension: ext,
            format: 'pdf',
            source: sourceOf(fsService, agent, target, info),
            output: { owner: agent.session.id, name: 'outline-layout.pdf' },
          }, exec.signal)
          const pdfBytes = await readFile(converted.path)
          layout = { pdf: parsePdf(pdfBytes), pageCount: converted.pageCount }
        } catch (error) {
          // A layout failure must not hide the structural facts that are still
          // available; the manifest says so instead of pretending to know pages.
          const model = scanStructure(bytes, ext)
          model.notes.push('分页失败：' + messageOf(error) + '。页数未知，按页渲染前请先整页渲染确认。')
          const decision = pageDecision(model)
          const advice = '分页不可用：先用 office_text 读文字并看 textChannel；需要看版面时按需渲染，并如实说明页码未能确认。'
          return {
            path: fsService.processPath(target),
            format: ext,
            family,
            pageCount: 0,
            pageCountSource: '',
            imagePages: [],
            textPages: [],
            advice,
            manifest: manifestOf(args.input, model, decision, advice),
            notes: model.notes,
          }
        }
      }
      const model = scanStructure(bytes, ext, layout)
      const decision = pageDecision(model)
      const advice = adviceOf(model, decision)
      return {
        path: fsService.processPath(target),
        format: ext,
        family: model.family,
        pageCount: model.pageCount,
        pageCountSource: model.pageCountSource,
        imagePages: decision.imagePages,
        textPages: decision.textPages,
        advice,
        manifest: manifestOf(args.input, model, decision, advice),
        notes: model.notes,
      }
    },
  }))

  ctx.tools.register(defineTool({
    name: 'office_text',
    description: [
      '把 Word 文档（doc/docx/odt）转成"每段一行、LF 换行"的 UTF-8 文本，并返回行数/字数/最长行。',
      '文本通道专门修掉了转换器用 CR 分隔段落导致整篇只有 1 行的问题，返回的文件可直接用 grep 按行检索。',
      '用 offset/limit 分段读（默认 1/60，最多 300 行）。textChannel=suspect 表示文本通道仍异常，必须停止按行检索并换读法。',
    ].join('\n'),
    parameters: {
      input: { type: 'string', required: true, description: 'Word 文档路径；相对路径按当前工作目录解析。' },
      offset: { type: 'integer', description: '起始行号（1 起）。默认 1。' },
      limit: { type: 'integer', description: '本次返回的最大行数。默认 60，最多 300。' },
    },
    output: {
      schema: {
        type: 'object',
        additionalProperties: false,
        properties: {
          path: { type: 'string', required: true },
          file: { type: 'string', required: true },
          lines: { type: 'integer', required: true },
          chars: { type: 'integer', required: true },
          maxLineLength: { type: 'integer', required: true },
          textChannel: { type: 'string', required: true },
          reason: { type: 'string', required: true },
          start: { type: 'integer', required: true },
          count: { type: 'integer', required: true },
          nextOffset: { type: 'integer', required: true },
          manifest: { type: 'string', required: true },
        },
      },
      render: (_args, value) => [{ type: 'text', text: value.manifest }],
    },
    async execute(args, exec) {
      const ext = extensionOf(args.input)
      if (!WORD_FAMILY.has(ext)) {
        throw new Error('office_text 只支持 Word 文档（doc/docx/odt）：.' + ext + ' 没有文本通道。PDF/表格/幻灯片请用 office_render 按页读，不要用 office_convert 的 txt 通道判断内容。')
      }
      const { fsService, agent, target, info } = await resolveInput(ctx, exec, args.input)
      const kit = officeService(ctx)
      let converted
      try {
        converted = await kit.convertTo({
          extension: ext,
          format: 'txt',
          source: sourceOf(fsService, agent, target, info),
          output: { owner: agent.session.id, name: 'text-source.txt' },
        }, exec.signal)
      } catch (error) {
        throw new Error('文本通道不可用：' + messageOf(error) + '。不要反复重试：改用 office_outline 的清单判断，或按页渲染。')
      }
      if (converted.bytes > MAX_TEXT_BYTES) {
        throw new Error('文本通道结果超过 ' + Math.round(MAX_TEXT_BYTES / 1048576) + 'MB，已停止：请分段读取并说明限制。')
      }
      const raw = await readFile(converted.path, 'utf8')
      const normalized = normalizeText(raw)
      // The official converter keeps its output in a private, Session-scoped
      // directory; the LF copy sits beside it and expires with that directory.
      // A failed copy must not cost the text channel: the slice below still
      // returns the document, and the manifest says so instead of linking a file.
      let file = join(dirname(converted.path), 'text-lf.txt')
      let fileNote = ''
      try {
        await writeFile(file, normalized.text.length === 0 ? '' : normalized.text + '\n', 'utf8')
      } catch (error) {
        file = ''
        fileNote = '（分段文件写入失败：' + messageOf(error) + '；请用 offset/limit 分段读，不要依赖 grep 路径）'
      }

      const offset = Math.max(1, Math.trunc(args.offset ?? 1))
      const limit = Math.max(1, Math.min(Math.trunc(args.limit ?? 60), 300))
      const all = normalized.text.length === 0 ? [] : normalized.text.split('\n')
      if (all.length > 0 && offset > all.length) {
        return {
          path: fsService.processPath(target),
          file,
          lines: normalized.lines,
          chars: normalized.chars,
          maxLineLength: normalized.maxLineLength,
          textChannel: normalized.textChannel,
          reason: normalized.reason,
          start: 0,
          count: 0,
          nextOffset: 0,
          manifest: 'office_text · ' + args.input + '\n已经到末尾：全文只有 ' + all.length + ' 行。' + (normalized.textChannel === 'suspect' ? '\n⚠ 文本通道异常：' + normalized.reason : '') + '\n分段文件：' + file,
        }
      }
      const start = Math.min(offset, Math.max(all.length, 1))
      const picked = []
      let used = 0
      for (let index = start - 1; index < all.length && picked.length < limit; index += 1) {
        const line = all[index]
        if (used + line.length > MAX_SLICE_CHARS && picked.length > 0) break
        picked.push(line)
        used += line.length + 1
      }
      const count = picked.length
      const nextOffset = start + count <= all.length ? start + count : 0
      const header = 'office_text · ' + args.input + '\n'
        + normalized.lines + ' 行 · ' + normalized.chars + ' 字 · 最长行 ' + normalized.maxLineLength + ' 字 · 文本通道 ' + normalized.textChannel + '\n'
        + (normalized.textChannel === 'suspect' ? '⚠ 文本通道异常：' + normalized.reason + '。停止按行检索，改用 offset/limit 分段或改用 office_render 整页阅读，不要反复重试。\n' : '')
        + (file === '' ? '分段文件：不可用' + fileNote + '\n' : '分段文件（LF 换行，可用 grep 按行检索）：' + file + '\n')
        + '第 ' + (count === 0 ? 0 : start) + '–' + (count === 0 ? 0 : start + count - 1) + ' 行' + (nextOffset > 0 ? '（下一段 offset=' + nextOffset + '）' : '（已到末尾）')
      const body = picked.map((line, index) => String(start + index).padStart(4, ' ') + '\t' + line).join('\n')
      return {
        path: fsService.processPath(target),
        file,
        lines: normalized.lines,
        chars: normalized.chars,
        maxLineLength: normalized.maxLineLength,
        textChannel: normalized.textChannel,
        reason: normalized.reason,
        start: count === 0 ? 0 : start,
        count,
        nextOffset,
        manifest: header + (count === 0 ? '\n（文档没有可读文本）' : '\n' + body),
      }
    },
  }))
}
