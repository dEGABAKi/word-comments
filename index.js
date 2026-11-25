/**
 * Native Word comments (.docx) as a standalone companion plugin.
 *
 * Deliberately independent from @copylee/dsh-office: it edits the document
 * package directly (word/comments.xml), so plugin upgrades cannot take the
 * feature away and no running Word/WPS instance is needed.
 */
import { spawn } from 'node:child_process'
import fs from 'node:fs'
import os from 'node:os'
import path from 'node:path'
import { fileURLToPath } from 'node:url'

export const name = 'word-comments'
export const inject = ['tools']

const SCRIPT = fileURLToPath(new URL('./py/comment_core.py', import.meta.url))
const RUN_TIMEOUT_MS = 120_000
const MAX_OUTPUT_BYTES = 4 * 1024 * 1024

/** Locate a Python that can import python-docx. */
function findPython() {
	const explicit = process.env.DSH_WORD_COMMENTS_PYTHON
	if (explicit && fs.existsSync(explicit)) return explicit

	const home = process.env.DSH_HOME || path.join(os.homedir(), '.dsh')
	const roots = path.join(home, 'dsh-runtimes')
	try {
		const runtimes = fs
			.readdirSync(roots, { withFileTypes: true })
			.filter((entry) => entry.isDirectory())
			.map((entry) => entry.name)
			// Prefer the bundled primary runtime when it exists.
			.sort((a, b) => Number(b === 'dsh-primary-runtime') - Number(a === 'dsh-primary-runtime'))
		for (const runtime of runtimes) {
			const candidate = path.join(
				roots,
				runtime,
				'dependencies',
				'python',
				process.platform === 'win32' ? 'python.exe' : path.join('bin', 'python3'),
			)
			if (fs.existsSync(candidate)) return candidate
		}
	} catch {
		// Fall through to PATH.
	}
	return process.platform === 'win32' ? 'python.exe' : 'python3'
}

function runCore(args, signal) {
	return new Promise((resolve, reject) => {
		const python = findPython()
		const child = spawn(python, [SCRIPT, ...args], {
			windowsHide: true,
			env: { ...process.env, PYTHONIOENCODING: 'utf-8', PYTHONUTF8: '1' },
		})

		let stdout = ''
		let stderr = ''
		let overflow = false
		let timedOut = false

		const collect = (buffer, chunk) => {
			const next = buffer + chunk.toString('utf8')
			if (next.length > MAX_OUTPUT_BYTES) {
				overflow = true
				return next.slice(0, MAX_OUTPUT_BYTES)
			}
			return next
		}

		child.stdout.on('data', (chunk) => {
			stdout = collect(stdout, chunk)
			if (overflow) child.kill()
		})
		child.stderr.on('data', (chunk) => {
			stderr = collect(stderr, chunk)
		})

		const timer = setTimeout(() => {
			timedOut = true
			child.kill()
		}, RUN_TIMEOUT_MS)

		const onAbort = () => child.kill()
		signal?.addEventListener('abort', onAbort, { once: true })

		const finish = () => {
			signal?.removeEventListener('abort', onAbort)
		}

		child.on('error', (error) => {
			clearTimeout(timer)
			finish()
			reject(error)
		})

		child.on('close', (code) => {
			clearTimeout(timer)
			finish()
			if (timedOut) {
				reject(new Error(`word comments helper timed out after ${RUN_TIMEOUT_MS / 1000}s`))
				return
			}
			if (overflow) {
				reject(new Error('word comments helper produced more output than expected'))
				return
			}
			resolve({ code, stdout, stderr })
		})
	})
}

function fail(message) {
	throw new Error(message)
}

function tail(text, limit = 1200) {
	const trimmed = String(text || '').trim()
	return trimmed.length > limit ? `…${trimmed.slice(-limit)}` : trimmed
}

function parseHelperOutput(run) {
	const text = run.stdout.trim()
	if (!text) {
		const detail = tail(run.stderr)
		fail(detail ? `word comments helper produced no result: ${detail}` : 'word comments helper produced no result')
	}
	try {
		return JSON.parse(text)
	} catch {
		fail(`word comments helper returned unreadable output: ${tail(text) || tail(run.stderr)}`)
	}
}

function resolveDocx(raw) {
	if (typeof raw !== 'string' || !raw.trim()) fail('path is required')
	const file = path.isAbsolute(raw) ? raw : path.resolve(process.cwd(), raw)
	if (!fs.existsSync(file)) fail(`no such file: ${file}`)
	if (path.extname(file).toLowerCase() !== '.docx') {
		fail(`${path.basename(file)} is not a .docx; convert it first (Word/WPS "Save as" .docx)`)
	}
	return file
}

function normalizeItems(items, author) {
	if (!Array.isArray(items) || items.length === 0) fail('comments is required for op=add')
	return items.map((item, index) => {
		const where = `comments[${index}]`
		if (!item || typeof item !== 'object') fail(`${where} must be an object`)
		const para = Number(item.para)
		if (!Number.isInteger(para) || para < 1) fail(`${where}.para must be an integer >= 1`)
		const text = typeof item.text === 'string' ? item.text : item.comment
		if (typeof text !== 'string' || !text.trim()) fail(`${where}.text must be a non-empty string`)
		return {
			para,
			expect: typeof item.expect === 'string' ? item.expect : undefined,
			find: typeof item.find === 'string' ? item.find : undefined,
			occurrence: Number.isInteger(item.occurrence) && item.occurrence > 0 ? item.occurrence : undefined,
			text,
			author: typeof item.author === 'string' ? item.author : undefined,
			initials: typeof item.initials === 'string' ? item.initials : undefined,
		}
	})
}

const output = {
	schema: {
		type: 'object',
		properties: {
			ok: { type: 'boolean' },
			op: { type: 'string' },
			path: { type: 'string' },
			source: { type: 'string' },
			copied: { type: 'boolean' },
			count: { type: 'integer' },
			comments: { type: 'array', items: { type: 'object', additionalProperties: true } },
			added: { type: 'array', items: { type: 'object', additionalProperties: true } },
			deleted: { type: 'array', items: { type: 'integer' } },
			remaining: { type: 'integer' },
		},
		additionalProperties: true,
	},
	render(args, value) {
		const result = value && typeof value === 'object' ? value : {}
		const lines = []
		if (result.copied && result.source && result.path) {
			const forced = args && typeof args === 'object' && args.copy === true
			lines.push(
				forced
					? `Wrote to a copy: ${result.path}`
					: `The document is open, so the change went to a copy: ${result.path}`,
			)
		}
		if (result.op === 'read') {
			const list = Array.isArray(result.comments) ? result.comments : []
			lines.push(`${list.length} comment${list.length === 1 ? '' : 's'} in ${result.path}`)
			for (const comment of list) {
				const anchor = comment.anchor_text ? ` on ${JSON.stringify(comment.anchor_text)}` : ''
				const para = comment.anchor_para ? ` ¶${comment.anchor_para}` : ''
				lines.push(`  #${comment.id}${para}${anchor} — ${comment.author || 'unknown'}: ${comment.text}`)
			}
		} else if (result.op === 'add') {
			const added = Array.isArray(result.added) ? result.added : []
			lines.push(`Added ${added.length} comment${added.length === 1 ? '' : 's'} to ${result.path}`)
			for (const item of added) {
				lines.push(`  #${item.id} ¶${item.para} on ${JSON.stringify(item.anchor_text || '')}`)
			}
		} else if (result.op === 'delete') {
			const deleted = Array.isArray(result.deleted) ? result.deleted : []
			lines.push(`Deleted ${deleted.length} comment${deleted.length === 1 ? '' : 's'} from ${result.path} (${result.remaining} left)`)
		}
		return [{ type: 'text', text: lines.join('\n') }]
	},
}

const tool = {
	name: 'word_comments',
	description: [
		'Add, list and remove native Word comments in a .docx file. Comments look and behave like ones a person inserted in Word or WPS Writer, and stay editable there.',
		'Works on the file, not on an open window: if the document is currently open in Word or WPS, the change is written to a sibling copy named "<name>-批注版.docx" and the reply gives that path, leaving the open document untouched. Pass copy=true to always work on a copy.',
		'Only .docx files; convert .doc first. Paragraph numbers are 1-based and count every paragraph in the document body, tables included, matching what office_read shows.',
	].join('\n'),
	parameters: {
		type: 'object',
		properties: {
			path: {
				type: 'string',
				description: 'Absolute path of the .docx file.',
			},
			op: {
				type: 'string',
				enum: ['read', 'add', 'delete'],
				description: 'read lists comments; add inserts them; delete removes them by id.',
			},
			comments: {
				type: 'array',
				description: 'Comments to insert when op is add. Each one anchors to one paragraph, optionally to the n-th occurrence of find inside it.',
				items: {
					type: 'object',
					properties: {
						para: {
							type: 'integer',
							description: '1-based paragraph number in the document body.',
						},
						text: {
							type: 'string',
							description: 'The comment text itself.',
						},
						expect: {
							type: 'string',
							description: 'Guard: the paragraph must start with this text, so a stale paragraph number fails instead of commenting the wrong place.',
						},
						find: {
							type: 'string',
							description: 'Anchor the comment on this text inside the paragraph instead of the whole paragraph.',
						},
						occurrence: {
							type: 'integer',
							description: 'Which occurrence of find to use (default 1).',
						},
						author: {
							type: 'string',
							description: 'Author for this comment, overriding the call-wide author.',
						},
						initials: {
							type: 'string',
							description: 'Author initials for this comment.',
						},
					},
					required: ['para', 'text'],
					additionalProperties: false,
				},
			},
			ids: {
				description: 'Comment ids to delete when op is delete, or "all".',
				oneOf: [
					{ type: 'array', items: { type: 'integer' } },
					{ type: 'string', enum: ['all'] },
				],
			},
			author: {
				type: 'string',
				description: 'Author shown in the comment balloon for inserted comments.',
			},
			copy: {
				type: 'boolean',
				description: 'true writes to the sibling copy even when the file is free; false edits in place and fails if the file is open. Omitted picks automatically.',
			},
		},
		required: ['path', 'op'],
		additionalProperties: false,
	},
	output,
	timeoutMs: RUN_TIMEOUT_MS + 10_000,
	isConcurrencySafe(args) {
		return args?.op === 'read'
	},
	async execute(args, exec) {
		const input = args && typeof args === 'object' ? args : {}
		const op = input.op
		if (op !== 'read' && op !== 'add' && op !== 'delete') fail('op must be read, add or delete')
		const file = resolveDocx(input.path)
		const copyFlag = input.copy === true ? '--copy' : input.copy === false ? '--no-copy' : undefined

		let run
		if (op === 'read') {
			run = await runCore(['read', file], exec?.signal)
		} else if (op === 'add') {
			const items = normalizeItems(input.comments, input.author)
			const payload = { items, author: typeof input.author === 'string' ? input.author : '' }
			const temp = path.join(
				os.tmpdir(),
				`word-comments-${process.pid}-${Date.now()}-${Math.random().toString(36).slice(2, 8)}.json`,
			)
			fs.writeFileSync(temp, JSON.stringify(payload), 'utf8')
			try {
				run = await runCore(['add', file, `@${temp}`, ...(copyFlag ? [copyFlag] : [])], exec?.signal)
			} finally {
				fs.rmSync(temp, { force: true })
			}
		} else {
			let ids
			if (input.ids === 'all') ids = 'all'
			else if (Array.isArray(input.ids) && input.ids.length > 0) {
				ids = input.ids.map((id) => {
					const n = Number(id)
					if (!Number.isInteger(n) || n < 0) fail('ids must be non-negative comment ids')
					return n
				})
			} else fail('ids is required for op=delete ("all" or a list of ids)')
			run = await runCore(['delete', file, String(ids), ...(copyFlag ? [copyFlag] : [])], exec?.signal)
		}

		const result = parseHelperOutput(run)
		if (result.ok === false) fail(result.error || 'word comments helper failed')
		if (run.code !== 0) fail(`word comments helper exited with code ${run.code}: ${tail(run.stderr)}`)
		return result
	},
}

export function apply(ctx) {
	ctx.effect(() => ctx.tools.register(tool), 'word-comments: tool word_comments')
}
