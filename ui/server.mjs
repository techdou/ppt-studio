#!/usr/bin/env node
/**
 * ppt-studio 本地查看器服务（零 npm 依赖，完全离线）。
 *
 * 用法:
 *   node server.mjs --project /abs/path/<项目名> [--port 55280]
 *
 * 职责:
 *   - 静态托管 ui/app/ 前端
 *   - 代理图片预览（项目 .qa-images/pages/N.png）
 *   - 调 scripts/export_images.py 触发 COM 重新导出
 *   - 调 scripts/page_text_tool.py 读写 .page 文本元素
 */
import { spawn } from "node:child_process";
import { createServer } from "node:http";
import { createReadStream, existsSync, mkdtempSync, readdirSync, statSync } from "node:fs";
import { readFile, rm, writeFile } from "node:fs/promises";
import { basename, join, normalize, resolve } from "node:path";
import { tmpdir } from "node:os";
import { fileURLToPath } from "node:url";
import { setTimeout as sleep } from "node:timers/promises";

const HERE = fileURLToPath(new URL(".", import.meta.url)); // ui/
const SKILL_ROOT = resolve(HERE, "..");
const SCRIPTS = join(SKILL_ROOT, "scripts");
const APP_DIR = join(HERE, "app");
const PYTHON = process.platform === "win32" ? "python" : "python3";

// ---------- 参数 ----------
function parseArgs(argv) {
  const opts = { project: null, port: 55280 };
  for (let i = 2; i < argv.length; i++) {
    if (argv[i] === "--project") opts.project = resolve(argv[++i]);
    else if (argv[i] === "--port") opts.port = Number(argv[++i]);
  }
  if (!opts.project) {
    console.error("缺少 --project <PPTD 项目目录>");
    process.exit(2);
  }
  return opts;
}
const opts = parseArgs(process.argv);

// ---------- 项目发现 ----------
function findDeck(projectDir) {
  if (!existsSync(projectDir)) throw new Error(`项目目录不存在: ${projectDir}`);
  const decks = readdirSyncSafe(projectDir).filter((f) => f.toLowerCase().endsWith(".pptd"));
  if (decks.length === 0) throw new Error(`项目目录下未找到 .pptd 清单: ${projectDir}`);
  decks.sort();
  if (decks.length > 1) console.warn(`[ppt-studio] 发现多个 .pptd，使用: ${decks[0]}`);
  return join(projectDir, decks[0]);
}

function readdirSyncSafe(dir) {
  try { return readdirSync(dir); } catch { return []; }
}

const DECK = findDeck(opts.project);
const QA_DIR = join(opts.project, ".qa-images");
const PAGES_PNG_DIR = join(QA_DIR, "pages");

// ---------- 子进程 ----------
function runPython(args, timeoutMs = 180_000) {
  return new Promise((res, rej) => {
    const p = spawn(PYTHON, args, { cwd: SCRIPTS, windowsHide: true });
    let out = "", err = "";
    const timer = setTimeout(() => { p.kill(); rej(new Error("python 调用超时")); }, timeoutMs);
    p.stdout.on("data", (d) => (out += d));
    p.stderr.on("data", (d) => (err += d));
    p.on("error", (e) => { clearTimeout(timer); rej(e); });
    p.on("close", (code) => { clearTimeout(timer); res({ code, out, err }); });
  });
}

async function getDeckInfo() {
  const r = await runPython([join(SCRIPTS, "page_text_tool.py"), "texts", DECK]);
  if (r.code !== 0) throw new Error(r.err.slice(0, 400) || "manifest 解析失败");
  const info = JSON.parse(r.out);
  const pngs = readdirSyncSafe(PAGES_PNG_DIR).filter((f) => /^\d+\.png$/i.test(f));
  return { ...info, pageCount: info.pages.length, hasQA: pngs.length > 0, qaDir: QA_DIR };
}

// ---------- HTTP 基础 ----------
const MIME = { ".html": "text/html; charset=utf-8", ".js": "text/javascript; charset=utf-8", ".css": "text/css; charset=utf-8", ".png": "image/png", ".jpg": "image/jpeg", ".svg": "image/svg+xml" };

function sendJSON(res, status, obj) {
  const body = JSON.stringify(obj);
  res.writeHead(status, { "Content-Type": "application/json; charset=utf-8", "Cache-Control": "no-store" });
  res.end(body);
}

function readFileSafe(path) { return existsSync(path) ? readFile(path) : null; }

// ---------- 路由 ----------
async function handleApi(req, res, url) {
  const route = url.pathname;

  if (route === "/api/project" && req.method === "GET") {
    try { return sendJSON(res, 200, { ...(await getDeckInfo()), project: opts.project, deck: basename(DECK) }); }
    catch (e) { return sendJSON(res, 500, { error: e.message }); }
  }

  const imgMatch = route.match(/^\/api\/image\/(\d+)\.png$/);
  if (imgMatch && req.method === "GET") {
    const file = join(PAGES_PNG_DIR, `${imgMatch[1]}.png`);
    if (!existsSync(file)) return sendJSON(res, 404, { error: "页面图未导出" });
    res.writeHead(200, { "Content-Type": "image/png", "Cache-Control": "no-store" });
    return createReadStream(file).pipe(res);
  }

  if (route === "/api/export" && req.method === "POST") {
    const before = pngMtimes();
    try {
      const r = await runPython([join(SCRIPTS, "export_images.py"), DECK, "--output", QA_DIR, "--force"], 300_000);
      if (r.code !== 0) return sendJSON(res, 500, { error: r.err.slice(0, 500) || "导出失败" });
      await sleep(200); // 文件系统 mtime 刷新缓冲
      const after = pngMtimes();
      const stale = before.size > 0 && [...after.entries()].every(([n, m]) => before.get(n)?.getTime() === m.getTime());
      return sendJSON(res, 200, { ok: true, stale,
        hint: stale ? "PNG 未见更新，可能存在残留 POWERPNT.EXE 进程返回旧副本，建议手动结束残留进程后重试" : null });
    } catch (e) { return sendJSON(res, 500, { error: e.message }); }
  }

  if (route === "/api/texts" && req.method === "GET") {
    const page = Number(url.searchParams.get("page") || 0);
    const args = [join(SCRIPTS, "page_text_tool.py"), "texts", DECK];
    if (page > 0) args.push("--page", String(page));
    try {
      const r = await runPython(args);
      if (r.code !== 0) return sendJSON(res, 500, { error: r.err.slice(0, 400) });
      return sendJSON(res, 200, JSON.parse(r.out));
    } catch (e) { return sendJSON(res, 500, { error: e.message }); }
  }

  if (route === "/api/update-text" && req.method === "POST") {
    let body = "";
    req.on("data", (d) => (body += d));
    req.on("end", async () => {
      try {
        const { page, elementId, text } = JSON.parse(body);
        if (!Number.isInteger(page) || !elementId || typeof text !== "string")
          return sendJSON(res, 400, { error: "参数不完整" });
        const tmp = mkdtempSync(join(tmpdir(), "ppt-studio-"));
        const tf = join(tmp, "text.txt");
        await writeFile(tf, text, "utf-8");
        try {
          const r = await runPython([join(SCRIPTS, "page_text_tool.py"), "update", DECK,
            "--page", String(page), "--element", String(elementId), "--text-file", tf]);
          if (r.code !== 0) return sendJSON(res, 500, { error: r.err.slice(0, 400) });
          return sendJSON(res, 200, JSON.parse(r.out));
        } finally {
          await rm(tmp, { recursive: true, force: true });
        }
      } catch (e) { return sendJSON(res, 500, { error: e.message }); }
    });
    return;
  }

  return sendJSON(res, 404, { error: "unknown api" });
}

function pngMtimes() {
  const m = new Map();
  for (const f of readdirSyncSafe(PAGES_PNG_DIR)) {
    if (/^\d+\.png$/i.test(f)) {
      try { m.set(f, statSync(join(PAGES_PNG_DIR, f)).mtime); } catch { /* ignore */ }
    }
  }
  return m;
}

// ---------- 静态文件 ----------
async function serveStatic(res, route) {
  let rel = route === "/" ? "index.html" : route.slice(1);
  const file = normalize(join(APP_DIR, rel));
  if (!file.startsWith(normalize(APP_DIR))) return sendJSON(res, 403, { error: "forbidden" });
  const dot = file.lastIndexOf(".");
  const ext = dot >= 0 ? file.slice(dot).toLowerCase() : "";
  const data = await readFileSafe(file);
  if (!data) return sendJSON(res, 404, { error: "not found" });
  res.writeHead(200, { "Content-Type": MIME[ext] || "application/octet-stream" });
  res.end(data);
}

const server = createServer((req, res) => {
  const url = new URL(req.url, `http://127.0.0.1:${opts.port}`);
  if (url.pathname.startsWith("/api/")) return void handleApi(req, res, url);
  serveStatic(res, url.pathname);
});

server.listen(opts.port, "127.0.0.1", () => {
  console.log(`[ppt-studio] 查看器已启动: http://127.0.0.1:${opts.port}/`);
  console.log(`[ppt-studio] 项目: ${DECK}`);
});
